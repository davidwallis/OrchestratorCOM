param()

<#
.SYNOPSIS
    Tests for the OrchestratorCOM module

.DESCRIPTION
    Contains various tests

.EXAMPLE
    Example: Invoke-Pester

.NOTES
    For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

.LINK
   https://github.com/davidwallis3101/OrchestratorCOM
#>

Function IsThisOnBuildServer {
    <#
    .SYNOPSIS
        Attempts to ensure that script signing and publishing only runs on an internal build server.
    .DESCRIPTION
        This validates the machine name that the build is running on to try and ensure
        that code is only published from a build server
    .EXAMPLE
        IsThisOnBuildServer
    .OUTPUTS
        Boolean
    .NOTES
        David Wallis
    #>
    [CmdLetBinding()]
    Param()
    Return $env:ComputerName -match "^BLD\d\d\d$"
}

# Get the base path
$ModuleBase = Split-Path -Parent $MyInvocation.MyCommand.Path

# Compatibility for when tests are located in .\Tests subdirectory
if ((Split-Path $ModuleBase -Leaf) -eq 'Tests') { $ModuleBase = Split-Path $ModuleBase -Parent }

# ChangeLog File
$changeLogPath = "$ModuleBase\CHANGELOG.md"

# Script under test
$sut = $ModuleBase + "\" + (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

# Construct the manifest file name
$ManifestFile = $ModuleBase + "\" + (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.tests\.ps1', '.psd1'

# Construct the module file name
$Module = $ModuleBase + "\" + (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.tests\.ps1', '.psm1'

#Construct the module name
$ModuleName = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.tests\.ps1'

$PSRepository = 'PSPrivateGallery'

#Import required dependencies
If (!(Get-PSRepository -Name $PSRepository -errorAction SilentlyContinue)) { write-error "Repository $PSRepository doesn't exist on this machine" }
Write-Host "Re-Installing any dependant modules"
# Add here

Write-Host "Loading any required snapins"
# Add here

#Remove existing copies of this module
Write-Host "Removing existing copies of $ModuleName"
Get-Module $ModuleName | Remove-Module -force

# Import the module
Write-Host "Installing module: $ModuleName"
import-module $module -Force -ErrorAction Stop

Describe "Module manifest ($ManifestFile)" -Tags "Script" {
    $Script:manifestVersion = $null
    $ManifestHash = Invoke-Expression (Get-Content $ManifestFile -Raw)

    It "Module - has a valid manifest file" {
        {
          $null = Test-ModuleManifest -Path $ManifestFile -ErrorAction Stop -WarningAction SilentlyContinue
        } | Should Not Throw
    }

    It "Manifest - has a valid root module" {
        $ManifestHash.RootModule | Should Be "$ModuleName.psm1"
    }

    It "Manifest - has a valid Description" {
        $ManifestHash.Description | Should Not BeNullOrEmpty
    }

    It "Manifest - specifies minimum powershell version" {
        $ManifestHash.PowerShellVersion | Should Not BeNullOrEmpty
    }

    It "Manifest - has a valid guid" {
        $ManifestHash.Guid | Should Be '92d43bb2-7f30-4c25-8e2c-531c569a14c3'
    }

    It "Manifest - has a valid version" {
        $Script:manifestVersion = $ManifestHash.ModuleVersion
        $ManifestHash.ModuleVersion -as [Version] | Should Not BeNullOrEmpty
    }

    It "Manifest - has a valid copyright" {
        $ManifestHash.Copyright | Should Not BeNullOrEmpty
    }

    It "Manifest - has a valid prefix" {
        $ManifestHash.Prefix | Should BeNullOrEmpty
    }

    It 'Manifest - has a valid project Uri' {
        $ManifestHash.ProjectUri | Should BeNullOrEmpty
    }

    It "Manifest - has a valid gallery tags without containing spaces" {
        foreach ($Tag in $ManifestHash.PrivateData.Values.tags)
        {
            $Tag -notmatch '\s' | Should Be $true
        }
    }

    It 'Manifest - contains required modules' {
        $ManifestHash.RequiredModules | Should BeNullOrEmpty
    }
}

Describe 'ChangeLog' -Tags "Script" {

    It "Module - has a valid changelog file" {
        Test-Path -LiteralPath $changeLogPath | Should Be $true
    }

    $script:changelogVersion = $null
    It "Module - has a valid version in the changelog" {
        foreach ($line in (Get-Content $changeLogPath -ErrorAction Stop))
        {
            if ($line -match "^\D*(?<Version>(\d+\.){1,3}\d+)")
            {
                $Script:changelogVersion = $matches.Version
                break
            }
        }
         $Script:changelogVersion| Should Not BeNullOrEmpty
         $Script:changelogVersion -as [Version]  | Should Not BeNullOrEmpty
    }

    if (IsThisOnBuildServer) {
        # Building on a build server so the change log should be correct.
        It "Module - changelog and manifest Major, Minor and Build versions are the same" {
            [Version]$changelogVersion = $Script:changelogVersion
            [Version]$manifestVersion = $Script:manifestVersion

            $changelogVersion.Major | Should be $manifestVersion.Major
            $changelogVersion.Minor | Should be $manifestVersion.Minor
            $changelogVersion.Build | Should be $manifestVersion.Build
        }
    }
    else
    {
        # Building on local machine so we dont want to be changing the changelog for each compile, only at publish time
        It "Module - changelog and manifest Major and Minor are the same" {
            [Version]$changelogVersion = $Script:changelogVersion
            [Version]$manifestVersion = $Script:manifestVersion

            $changelogVersion.Major | Should be $manifestVersion.Major
            $changelogVersion.Minor | Should be $manifestVersion.Minor
        }
    }
}

Describe 'Function help' -Tags "Script" {
    $Commands = (get-command -Module $ModuleName -CommandType Cmdlet, Function, Workflow)
    ForEach ($Command in $Commands)
    {
        $Help = Get-Help -Name $command.Name -ErrorAction SilentlyContinue

        if ($Help.Synopsis -like '*`[`<CommonParameters`>`]*') {
            $Help = Get-Help $command.Name -ErrorAction SilentlyContinue
        }

        # Get the notes fields
        $Notes = ($Help.alertSet.alert.text -split '\n')

        # Parse the function using AST
        $AST = [System.Management.Automation.Language.Parser]::ParseInput((Get-Content function:$Command), [ref]$null, [ref]$null)

        Context "Function $Command - Validate Help"{

            It "Function - Synopsis should not be auto generated" { $Help.Synopsis | Should Not BeLike '*`[`<CommonParameters`>`]*' }

            It "Function - has a valid Synopsis"{ $help.Synopsis | Should not BeNullOrEmpty }

            It "Function - has a valid Description"{ $help.Description | Should not BeNullOrEmpty }

            # If function supports shouldprocess then test that the confirm impact is set to Medium or High
            if ($Command.Parameters['Whatif']) {
                $metadata = [System.Management.Automation.CommandMetadata]$command

                It "Function - has ConfirmImpact of Medium or High configured" {
                    $metadata.ConfirmImpact | Should Match 'Medium|High'
                }
            }

            #It "Function - has a valid Link"{ $help.Link | Should be BeNullOrEmpty }

            It "Function - Notes" { $Notes[0].trim() | Should not be NullOrEmpty }

            # Get the parameters declared in the Comment Based Help
            $RiskMitigationParameters = 'Whatif', 'Confirm'
            $HelpParameters = $help.parameters.parameter | Where-Object name -NotIn $RiskMitigationParameters

            # Validate Help start at the beginning of the line
            #$FunctionContent = Get-Content function:$Command
            # It "Function - Comment based help starts at the beginning of the line"{
            #     $Pattern = ".Synopsis"
            #     ($FunctionContent -split '\r\n' |
            #         select-string $Pattern).line -match "^$Pattern" | Should Be $true
            # }

            # Get the parameters declared in the AST PARAM() Block
            $ASTParameters = $AST.ParamBlock.Parameters.Name.variablepath.userpath

            It "Parameter - Compare Count Help/AST" {
                $HelpParameters.name.count -eq $ASTParameters.count | Should Be $true
            }

            # Check Parameters have descriptions
            $HelpParameters | ForEach-Object {
                $Parameter = $Command.ParameterSets.Parameters | Where-Object Name -eq $_.Name

                # Stops functions with no params erroring.
                if ($_.Name) {
                    It "Parameter - $($_.Name) contains a valid parameter description in the comment based help" { $_.description | Should not BeNullOrEmpty }

                    # Parameter type in Help should match code
                    It "Parameter - $($_.Name) has correct parameter type" {
                        $codeType = $Parameter.ParameterType.Name
                        $helpType = if ($_.parameterValue) { $_.parameterValue.Trim() }
                        $helpType | Should be $codeType
                    }

                    It "Parameter - $($_.Name) has correct mandatory value" {
                        $codeMandatory = $Parameter.IsMandatory.toString()
                        $_.Required | Should Be $codeMandatory
                    }
                }
            }

            # Using Abstract Syntax Tree (AST), retrieve the content of the PARAM() block and split on the carriage return character
            $ParamText = $AST.ParamBlock.extent.text -split '\r\n'
            $ParamText = $ParamText.trim()
            $ParamTextSeparator = $ParamText | select-string ',$' #line that finish by a ','

            if ($ParamTextSeparator)
            {
                Foreach ($ParamLine in $ParamTextSeparator.linenumber)
                {
                    it "Parameter - Separated by space (Line $ParamLine)"{ $ParamText[$ParamLine] -match '^$|\s+' | Should Be $true }
                }
            }

            # Each function should have at least one code examples present
            it "Examples - Count should be greater than 0"{ $Help.Examples.Example.Code.Count | Should BeGreaterthan 0 }

            # Each example should have a description
            foreach ($Example in $Help.Examples.Example)
            {
                it "Examples - Description on $($Example.Title)" { $Example.Remarks | Should not BeNullOrEmpty }
            }
        }
    }
}

Describe 'Style rules' -Tags "Script" {

    #$pesterRoot = (Get-Module Pester).ModuleBase
    $files = @( Get-ChildItem $ModuleBase -Include *.ps1,*.psm1 -Recurse )

    It 'Style - Source files contain no trailing whitespace' {
        $badLines = @(
            foreach ($file in $files)
            {
                $lines = [System.IO.File]::ReadAllLines($file.FullName)
                $lineCount = $lines.Count

                for ($i = 0; $i -lt $lineCount; $i++)
                {
                    if ($lines[$i] -match '\s+$')
                    {
                        'File: {0}, Error on line: [{1}]' -f $file.FullName, ($i + 1)
                    }
                }
            }
        )

        if ($badLines.Count -gt 0)
        {
            throw "The following $($badLines.Count) lines contain trailing whitespace: `r`n`r`n$($badLines -join "`r`n")"
        }
    }

    # Check for tab's being used for indentation
    It 'Style - Uses spaces for indentation, not tabs' {
        $totalTabsCount = 0
        $files | ForEach-Object {
            $fileName = $_.FullName
            #$tabStrings = (Get-Content $_.FullName -Raw) | Select-String "`t" |  ForEach-Object {
            Get-Content $_.FullName -Raw | Select-String "`t" |  ForEach-Object {
                Write-Warning "There are tab(s) in $fileName. Please convert to space indentation'."
                $totalTabsCount++
            }
        }
        $totalTabsCount | Should Be 0
    }

    # To improve consistency across multiple environments and editors each text file is required to end with a new line.
    It 'Style - Source Files all end with a newline' {
        $badFiles = @(
            foreach ($file in $files)
            {
                $string = [System.IO.File]::ReadAllText($file.FullName)
                if ($string.Length -gt 0 -and $string[-1] -ne "`n")
                {
                    $file.FullName
                }
            }
        )

        if ($badFiles.Count -gt 0)
        {
            throw "The following files do not end with a newline: `r`n`r`n$($badFiles -join "`r`n")"
        }
    }
}

Describe 'PowerShell Script Analyzer' -Tags "Script" {
    # PSScriptAnalyzer requires PowerShell 5.0 or higher
    if ($PSVersionTable.PSVersion.Major -ge 5)
    {
        Context 'PSScriptAnalyzer' {
            It 'passes Invoke-ScriptAnalyzer' {

                # Perform PSScriptAnalyzer scan.
                # Using ErrorAction SilentlyContinue not to cause it to fail due to parse errors caused by unresolved resources.
                # Code may try to import different modules which may not be present on the machine and PSScriptAnalyzer throws parse exceptions even though examples are valid.
                # Errors will still be returned as expected.

                #$PSScriptAnalyzerErrors = Invoke-ScriptAnalyzer -path  -Recurse -ErrorAction SilentlyContinue <# -Severity Error #>

                # Scan all files, excluding build sript
                # $PSScriptAnalyzerErrors = Get-ChildItem -Path $ModuleBase -Exclude "*.psake.ps1" -Recurse | ForEach-Object {
                #     Invoke-ScriptAnalyzer -path $_ -Severity Warning <# -Severity Error #> -ErrorAction SilentlyContinue -Verbose
                # }

                # scan only module files with recursion
                # $PSScriptAnalyzerErrors = Get-ChildItem -Path $ModuleBase -Include "*.psm1" -Recurse | ForEach-Object {
                #     Invoke-ScriptAnalyzer -path $_ -Severity Warning -ErrorAction SilentlyContinue -Verbose
                # }

                # scan only module files without recursion
                $PSScriptAnalyzerErrors = Get-ChildItem -Path $ModuleBase -Include "*.psm1" | ForEach-Object {
                    Invoke-ScriptAnalyzer -path $_ -Severity Warning -ErrorAction SilentlyContinue -Verbose
                }

                if ($PSScriptAnalyzerErrors -ne $null) {
                    Write-Warning -Message 'There are PSScriptAnalyzer errors that need to be fixed:'
                    @($PSScriptAnalyzerErrors).Foreach( { Write-Warning -Message "$($_.Scriptname) (Line $($_.Line)): $($_.Message)" } )
                    $PSScriptAnalyzerErrors |format-table
                    @($PSScriptAnalyzerErrors).Count | Should Be 0
                }
            }
        }
    }
    else
    {
        write-warning "Unable to run PowerShell Script Analyzer due to Powershell version"
    }
}

<#

https://github.com/pester/Pester/issues/311

http://bentaylor.work/2017/01/mocking-new-object-in-pester-with-powershell-classes/

http://scottmuc.com/testing-powershell-code-that-talks-to-clr-objects/
#>
