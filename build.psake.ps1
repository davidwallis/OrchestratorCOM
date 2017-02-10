##############################################################################
# PSAKE SCRIPT FOR MODULE BUILD & PUBLISH TO THE INTERNAL PSPRIVATEGALLERY
##############################################################################
#
# Requirements: PSake.  If you don't have this module installed use the following
# command to install it:
#
# PS C:\> Install-Module PSake -Scope CurrentUser
#
##############################################################################
# This is a PSake script that supports the following tasks:
# clean, build, test and publish.  The default task is build.
#
# The publish task uses the Publish-Module command to publish
# to the internal PowerShell Gallery
#
# The test task invokes Pester to run any Pester tests in your
# workspace folder. Name your test scripts <TestName>.Tests.ps1
# and Pester will find and run the tests contained in the files.
#
# You can run this build script directly using the invoke-psake
# command which will execute the build task.  This task "builds"
# a temporary folder from which the module can be published.
#
# PS C:\> invoke-psake build.ps1
#
# You can run your Pester tests (if any) by running the following command.
#
# PS C:\> invoke-psake build.ps1 -taskList test
#
# You can execute the publish task with the following command. Note that
# the publish task will run the test task first. The Pester tests must pass
# before the publish task will run.
#
# PS C:\> invoke-psake build.ps1 -taskList publish
#
# This command should only be run from a TFS Build server

###############################################################################
# Customize these properties for your module.
# PSake makes variables declared here available in other scriptblocks
###############################################################################

Properties {
    Write-Host "Using PowerShell Version: $($PSversionTable.PSversion.ToString())"
    $ManifestPath = (Get-Item $PSScriptRoot\*.psd1)[0]
    # The name of your module should match the basename of the PSD1 file.
    $ModuleName = $ManifestPath.BaseName

    Write-Host "Building module: $ModuleName"

    # Path to the release notes file.
    # $ReleaseNotesPath = "$PSScriptRoot\CHANGELOG.md"

    # The directory used to publish the module from.  If you are using Git, the
    # $PublishDir should be ignored if it is under the workspace directory.
    $PublishDir = "$PSScriptRoot\Release\$ModuleName"

    # The directory is used to publish the test results to.
    $TestResultsDir = "$PSScriptRoot\TestResults"

    # The following items will not be copied to the $PublishDir.
    # Add items that should not be published with the module.
    $Exclude = @(
        'Release',
        'TestResults',
        'Tests',
        '.git*',
        '.vscode',
        (Split-Path $PSCommandPath -Leaf)
    )

    # Name of the repository you wish to publish to
    $PublishRepository = 'PSPrivateGallery'
    $galleryServerUri = 'http://internalGallery:8080/api/v2'

    # NuGet API key for the gallery.
    $NuGetApiKey = 'myapikeywashere'

    # Stop pester outputing blank lines see issue #450
    $ProgressPreference = 'SilentlyContinue'

    if ($ENV:BUILD_BUILDNUMBER) {
        $BuildNumber = "$ENV:BUILD_BUILDNUMBER"
    } else {
        $BuildNumber = 0
    }
}

Task ScriptSigning {
    # Sign the scripts if we are running on a build server
    If ((IsThisOnBuildServer) -eq $False) {Write-Error "This step should be only run from a build server, stopping." -ErrorAction Stop}

    Write-Host "Signing scripts"
    #  We can only sign .ps1, .psm1, .psd1, and .ps1xml files
    Get-ChildItem ("{0}\*.ps*"-f $PublishDir) | Sign-Script
}

###############################################################################
# Customize these tasks for performing operations before and/or after publish.
###############################################################################

# Executes before src is copied to publish dir
Task PreCopySource {
}

# Executes after src is copied to publish dir
Task PostCopySource {
}

# Executes before publishing occurs.
Task PrePublish {
}

# Executes after publishing occurs.
Task PostPublish {
}

###############################################################################
# Core task implementations
###############################################################################
Task default -depends Build

Task Init -requiredVariables PublishDir {

    # Check if we are running on a tfs build server
    If (IsThisOnBuildServer) {
        $agent = $($env:AGENT_NAME)
    }
    else {
         $agent = "N/A"
    }

    Write-Host ("Build Running on: {0}, Build Agent: {1}" -f $($env:ComputerName), $agent)

    if (!(Test-Path $PublishDir)) {
        Write-Host "Creating Publish folder ($PublishDir)"
        $null = New-Item $PublishDir -ItemType Directory
    }

    if (!(Test-Path $TestResultsDir)) {
        Write-Host "Creating TestResults folder ($TestResultsDir)"
        $null = New-Item $TestResultsDir -ItemType Directory
    }

    If (!(Get-PSRepository -Name $PublishRepository -errorAction SilentlyContinue)) {
    Write-Host "Repository $PublishRepository doesn't exist, Adding"
    Register-PSRepository `
        -Name $PublishRepository `
        -SourceLocation $galleryServerUri `
        -InstallationPolicy Trusted `
        -PackageManagementProvider NuGet
    }
}

Task Clean -depends Init -requiredVariables PublishDir {
    # Sanity check the dir we are about to "clean".  If $PublishDir were to
    # inadvertently get set to $null, the Remove-Item commmand removes the
    # contents of \*.  That's a bad day.  Ask me how I know?  :-(
    if ($PublishDir.Contains($PSScriptRoot)) {
        Write-Host (' Cleaning publish directory "{0}".' -f $PublishDir)
        Remove-Item $PublishDir\* -Recurse -Force
    }

    if ($TestResultsDir.Contains($PSScriptRoot)) {
        Write-Host (' Cleaning test results directory "{0}".' -f $TestResultsDir)
        Remove-Item $TestResultsDir\* -Recurse -Force
    }
}

Task Build -depends Clean -requiredVariables PublishDir, Exclude, ModuleName {

    If (IsThisOnBuildServer) {
        Step-ModuleVersion -Path $ManifestPath -Verbose -Increment None -RevisionNumber $BuildNumber
    }
    else {
        # Incrementing Build Version
        Step-ModuleVersion -Path $ManifestPath -Verbose -Increment Build

       $Version = Get-ModuleVersion -Path $ManifestPath
       "$ModuleName Version: $Version"
    }

    Copy-Item $PSScriptRoot\* -Destination $PublishDir -Recurse -Exclude $Exclude

    # Get contents of the ReleaseNotes file and update the copied module manifest file
    # with the release notes.
    # DO NOT USE UNTIL UPDATE-MODULEMANIFEST IS FIXED - HORRIBLY BROKEN RIGHT NOW.
    # if ($ReleaseNotesPath) {
    #      $releaseNotes = @(Get-Content $ReleaseNotesPath)
    #      Update-ModuleManifest -Path $PublishDir\${ModuleName}.psd1 -ReleaseNotes $releaseNotes
    # }
}

Task Test -depends TestImpl {
}

Task TestImpl -depends Build {
     # Run pester tests
     Import-Module Pester

     # Display pester version being used to run tests
     Write-Host ("Testing with Pester version {0}" -f (get-module pester).Version.ToString())

     $invokePesterParams = @{
         OutputFile = "$TestResultsDir/Test-Pester-$ModuleName-Script.XML";
         OutputFormat = 'NUnitXml';
         Strict = $true;
         PassThru = $true;
         Verbose = $false;
         Show = "Fails";
     }

    $testResult = Invoke-Pester @invokePesterParams -Script @{ Path = "$PSScriptRoot\tests\$ModuleName.tests.ps1"; }

    if ($testResult.FailedCount -gt 0) {
        Write-Error ('Failed "{0}" unit tests.' -f $testResult.FailedCount);
    }
}

Task Publish -depends Test, ScriptSigning, PrePublish, PublishImpl, PostPublish {
}

Task PublishImpl -depends Test -requiredVariables PublishDir {

    $publishParams = @{
        Path        = $PublishDir
        NuGetApiKey = $NuGetApiKey
        Repository  = $PublishRepository
    }

    # Consider not using -ReleaseNotes parameter when Update-ModuleManifest has been fixed.
    if ($ReleaseNotesPath) {
        $publishParams['ReleaseNotes'] = @(Get-Content $ReleaseNotesPath)
    }

    $Version = Get-ModuleVersion -Path $ManifestPath
    "Publishing module ($ModuleName Version: $Version) to gallery: $PublishRepository"
    Publish-Module @publishParams
}

Task ? -description 'Lists the available tasks' {
    "Available tasks:"
    $psake.context.Peek().tasks.Keys | Sort
}

###############################################################################
# Helper functions
###############################################################################
function Step-ModuleVersion {
    [CmdletBinding()]
    param(
        # Specifies a path a valid Module Manifest file.
        [Parameter(Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Path to one or more locations.")]
        [Alias("PSPath")]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string[]]$Path,

        # Version section to step
        [Parameter(Position=1)]
        [ValidateSet("Major", "Minor", "Build", "Revision", "None")]
        [Alias("Type")]
        [string]$Increment = "Build",

        [Parameter(Position=2)]
        [int]$RevisionNumber
    )

    Begin
    {
        if (-not $PSBoundParameters.ContainsKey("Path")) {
            $Path = (Get-Item $PWD\*.psd1)[0]
        }
    }

    Process
    {
        foreach ($file in $Path)
        {
            $manifest = Import-PowerShellDataFile -Path $file
            $newVersion = Step-Version `
                -Version $manifest.ModuleVersion `
                -Increment $Increment `
                -RevisionNumber $RevisionNumber

            $manifest.Remove("ModuleVersion")
            $manifest.FunctionsToExport = $manifest.FunctionsToExport | ForEach-Object {$_}
            $manifest.NestedModules = $manifest.NestedModules | ForEach-Object {$_}
            $manifest.RequiredModules = $manifest.RequiredModules | ForEach-Object {$_}
            $manifest.ModuleList = $manifest.ModuleList | ForEach-Object {$_}

            if ($manifest.ContainsKey("PrivateData") -and $manifest.PrivateData.ContainsKey("PSData")) {
                foreach ($node in $manifest.PrivateData["PSData"].GetEnumerator()) {
                    $key = $node.Key
                    if ($node.Value.GetType().Name -eq "Object[]") {
                        $value = $node.Value | ForEach-Object {$_}
                    }
                    else {
                        $value = $node.Value
                    }

                    $manifest[$key] = $value
                }
                $manifest.Remove("PrivateData")
            }
            New-ModuleManifest -Path $file -ModuleVersion $newVersion @manifest
        }
    }
}

function Get-ModuleVersion {
    [CmdletBinding()]
    param(
        # Specifies a path a valid Module Manifest file.
        [Parameter(Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Path to one or more locations.")]
        [Alias("PSPath")]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string[]]$Path
    )

    Begin {
        if (-not $PSBoundParameters.ContainsKey("Path")) {
            $Path = (Get-Item $PWD\*.psd1)[0]
        }
    }

    Process {
        foreach ($file in $Path) {
            $manifest = Import-PowerShellDataFile -Path $file
            return $manifest.ModuleVersion
        }
    }
}

function Step-Version {
    Param (
    [CmdletBinding()]
    [OutputType([String])]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [String]$Version,

        [Parameter(Mandatory=$true, Position=1)]
        [ValidateSet("Major", "Minor", "Build", "Revision", "None")]
        [Alias("By")]
        [string]$Increment,

        [Parameter(Position=2)]
        [string]$Revision = 0,

        [Parameter(Position=3)]
        [int]$RevisionNumber
    )

    Process {
        $currentVersion = [version]$Version

        [int]$major = $currentVersion.Major
        [int]$minor = $currentVersion.Minor
        [int]$build = $currentVersion.Build
        [int]$revision = $currentVersion.Revision

        if($Major -lt 0) { $Major = 0 }
        if($Minor -lt 0) { $Minor = 0 }
        if($Build -lt 0) { $Build = 0 }
        if($Revision -lt 0) { $Revision = 0 }

        switch($Increment) {
            "Major" {
                $major++
            }
            "Minor" {
                 $Minor++
            }
            "Build" {
                 $Build++
            }
            "Revision" {
                 $Revision++
            }
            "None" {
                # Append revision number from param instead.
                $Revision  = $RevisionNumber
            }
        }

        $newVersion = New-Object Version -ArgumentList $major, $minor, $Build, $Revision
        Write-Output -InputObject $newVersion.ToString()
    }
}

function Sign-Script {

    <#
    .SYNOPSIS
        Script for signing scripts during the build process
    .DESCRIPTION
        This script is called after the pester tests are successfuly completed and then signs the code
    .PARAMETER Path
        Provide a path to the file you want to sign
    .EXAMPLE
        C:\PS> SignScript c:\temp\test.ps1
    .NOTES
        David Wallis
    #>

    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true,ValueFromPipeLine=$true,ValueFromPipeLineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('FullName')]
        [string[]]$Path
    )

    Begin {
        $cert =(Get-Childitem cert:\CurrentUser\My -CodeSigningCert)
        If ($null -eq $cert) { write-error "No code signing certificate found" -ErrorAction Stop}

        If($cert[0].NotAfter -le (get-date)) {
            Write-error "Code signing certificate expired"
        } else {
            write-verbose "Certificate valid until: $($cert.NotAfter)"
        }
    }
    Process{
        Write-verbose "Signing script $($Path)"
        $AuthenticodeSignature = Set-AuthenticodeSignature $Path $cert[0] -ErrorAction stop
    }
}

Function IsThisOnBuildServer {
    <#
    .SYNOPSIS
        Attempts to ensure that script signing and publishing only runs on a build server.
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
    Return $env:ComputerName -match "^(PLL|VAL|CRL)WIN(DV|QA|LV|CS)BLD\d\d\d$"
}
