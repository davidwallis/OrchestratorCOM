# TODO by default make cmdlets such as get-orchestratorIntegrationPack get all 'IntegrationPacks' and allow filters to limit the name or guid.

# Use 32 bit powershell

# Store connection information
$script:connectionHandle = $null
$script:oisMgr = $null
$script:DatabaseServer = $null
$script:DatabaseName = $null

$ResourceType = @{}
$ResourceType.Variable = "{2E88BB5A-62F9-482E-84B0-4D963C987231}"
$ResourceType.Counter =  "{0BABBCF6-C702-4F02-9BA6-BAB75983A06A}"
$ResourceType.Schedule = "{4386DA28-C311-4A2B-8C47-3C7BB9D66B51}"
$ResourceType.Computer = "{162204B6-7F54-4CB9-A678-B94A6510BD0C}"
$ResourceType.Satellite  = "{155D5068-BFF4-4054-AD01-19403371FAD2}"

<#
* WAIT_OBJECTTYPE              = '{B40FDFBD-6E5F-44F0-9AA6-6469B0A35710}"
* CUSTOM_START_OBJECTTYPE      = '{6C576F3D-E927-417A-B145-5D3EFF9C995F}"
* LINK_OBJECTTYPE              = '{7A65BD17-9532-4D07-A6DA-E0F89FA0203E}"
* COUNTER_GET_OBJECTTYPE       = '{4E753C05-1A1F-4350-B572-09AE196AB593}"
* COUNTER_SET_OBJECTTYPE       = '{D2259B53-4C86-4A58-B40B-7493FC182E02}"
* COUNTER_MONITOR_OBJECTTYPE   = '{15A1CEB4-16D8-4AD7-A54B-8162A309439C}"
* CHECK_SCHEDULE_OBJECTTYPE    = '{0B807C4B-41C3-4517-B24E-7D98F016AD1C}"
* TRIGGER_POLICY_OBJECTTYPE    = '{9C1BF9B4-515A-4FD2-A753-87D235D8BA1F}"
* NOTEOBJECT_GUID              = '{AB1D2E56-3842-4184-A9AF-DFBB99115D26}"
* JUNCTION_TYPE                = '{1C5F9236-92E0-4795-8CAA-1669B7643607}"
#>

$ResourceFolderRoot = @{}
$ResourceFolderRoot.Runbooks = "{00000000-0000-0000-0000-000000000000}"
$ResourceFolderRoot.Computers = "{00000000-0000-0000-0000-000000000001}"
$ResourceFolderRoot.Reporting = "{00000000-0000-0000-0000-000000000002}"
$ResourceFolderRoot.RunbookServers = "{00000000-0000-0000-0000-000000000003}"
$ResourceFolderRoot.Counters = "{00000000-0000-0000-0000-000000000004}"
$ResourceFolderRoot.Variables = "{00000000-0000-0000-0000-000000000005}"
$ResourceFolderRoot.Schedules = "{00000000-0000-0000-0000-000000000006}"
$ResourceFolderRoot.Satellites = "{00000000-0000-0000-0000-000000000007}"

################################################ Public Functions ################################################

Function Connect-OrchestratorComInterface {
    <#
        .SYNOPSIS
            Connect-OrchestratorComInterface

        .DESCRIPTION
            Connects to the Orchestrator COM interface

        .PARAMETER Credential
            The credentials to be used when authenticating

        .EXAMPLE
            PS C:\> Connect-OrchestratorComInterface -Credential (Get-Credential domain\user)

        .INPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    #[OutputType([System.Int32])]
    param(
        [Parameter(Position=0)]
        [ValidateNotNullOrEmpty()]
        # due to ps script analyzer 'must be of type PSCredential. For PowerShell 4.0 and earlier, please define a credential transformation attribute'
        #[System.Net.NetworkCredential]
        # [System.Management.Automation.Credential()]$Credential = (Get-Credential)

        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = (Get-Credential)
    )

    Process {
        try {
            If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME)) {

                # Insert the domain name into the username if present
                if ($Credential.Domain) {
                    $username = "$($Credential.Domain)\"
                }
                $username += $credential.UserName

                Write-verbose ("Connecting to COM using Username: {0}" -f $Username)
                if ($null -eq $script:oisMgr) { $script:oisMgr = new-object -com OpalisManagementService.OpalisManager}

                # Get password from credential object, had to add this after psscriptanalyser complained about the use of
                # [System.Net.NetworkCredential] as the credential type.
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
                $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

                $oHandle = New-Object Object
                $handle = New-Object Runtime.InteropServices.VariantWrapper($oHandle)

                #$oisMgr.Connect($Username, $Credential.Password, [ref]$handle)
                $oisMgr.Connect($Username, $password, [ref]$handle)
                $script:connectionHandle = $handle

                Write-Verbose ("Got connection handle: {0}" -f $handle)
            }
        }
        catch {
            # NB the Variant wrapper use above doesn't appear to work with WMF5 installed!
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Disconnect-OrchestratorComInterface {
    <#
        .SYNOPSIS
            Disconnect-OrchestratorComInterface

        .DESCRIPTION
            Disconnects from the Orchestrator COM interface

        .EXAMPLE
            PS C:\> Disconnect-OrchestratorComInterface

        .INPUTS
            System.Int32

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    #[OutputType([System.Int32])]
    param(
    )

    Process {
        try {
            If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME)) {
                # Second string value purpose??
                $oisMgr.Disconnect($script:connectionHandle, "");
                write-verbose ("Disconnected handle {0}" -f $script:connectionHandle)
                $script:connectionHandle = $null
                $script:DatabaseServer = $null
                $script:DatabaseName = $null
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
    End {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oisMgr) | Out-Null
        $oisMgr = $null
    }
}

Function Get-OrchestratorIntegrationPacks {
    <#
        .SYNOPSIS
            Get-OrchestratorIntegrationPacks

        .DESCRIPTION
            Gets Orchestrator Integration Packs

        .EXAMPLE
            PS C:\> Get-OrchestratorIntegrationPacks

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        # [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        # [ValidateNotNullOrEmpty()]
        # [Int]$Handle
        # [Parameter(Position=0)]
        # [ValidateNotNullOrEmpty()]
        # [System.Net.NetworkCredential]$Credential = (Get-Credential)
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }
            If ($PSCmdlet.ShouldProcess("")) {
                $integrationPackArray = @() # array to store the results

                $oIPs = New-Object object
                $integrationPacks = New-Object Runtime.InteropServices.VariantWrapper($oIPs)
                $oismgr.GetIntegrationPacks($script:connectionHandle, [ref]$integrationPacks)

                $integrationPackNodes = ([XML]$integrationPacks).SelectNodes("//IntegrationPack")

                foreach ($integrationPackNode in $integrationPackNodes) {
                    $integrationPackArray += [PSCustomObject]@{
                            'Name'=[string]$integrationPackNode.Name.InnerText;
                            'Description'=[string]$integrationPackNode.Description.InnerText;
                            'Version'=[string]$integrationPackNode.Version.InnerText;
                            'Library'=[string]$integrationPackNode.Library.InnerText;
                            <#'ProductName'=[string]$integrationPackNode.ProductName.InnerText;#>
                            <#'ProductID'=[string]$integrationPackNode.ProductID.InnerText;#>
                            'Guid'=[Guid]$integrationPackNode.UniqueID.InnerText
                    }
                }
                return $integrationPackArray
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }

}

Function Get-OrchestratorPoliciesWithoutImages {
    <#
        .SYNOPSIS
            Get-OrchestratorPoliciesWithoutImages

        .DESCRIPTION
            Gets Orchestrator Policies Without Imagess

        .EXAMPLE
            PS C:\> Get-OrchestratorPoliciesWithoutImages

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param()

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }
            If ($PSCmdlet.ShouldProcess("")) {
                $oPolicies = New-Object object
                $policies = New-Object Runtime.InteropServices.VariantWrapper($oPolicies)
                $oismgr.FindPoliciesWithoutImages($script:connectionHandle, [ref]$policies)
                # TODO output array, look to see what delim to split on - New line or } ?
                Write-Output $policies
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorPolicyRunningState {
    <#
        .SYNOPSIS
            Get-OrchestratorPolicyRunningState

        .DESCRIPTION
            Returns whether a policy is running or not

        .PARAMETER Policy
            The policy guid

        .EXAMPLE
            PS C:\> Get-OrchestratorPolicyRunningState

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [GUID[]]$Policy
    )

    Process {
        foreach ($pol in $Policy) {
            try {
                if (!$script:connectionHandle) {
                    Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
                }

                If ($PSCmdlet.ShouldProcess("")) {
                    # IsPolicyRunning (lHandle As Long, bstrPolicyID As String)
                    $oismgr.IsPolicyRunning($script:connectionHandle, $pol.ToString("B"))
                }
            }
            catch {
                Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
            }
        }
    }
}

Function Get-OrchestratorCheckOutStatus {
    <#
        .SYNOPSIS
            Get-OrchestratorCheckOutStatus

        .DESCRIPTION
            Returns whether a policy is checked out or not

        .PARAMETER  RunbookGUID
            The runbook guid

        .EXAMPLE
            PS C:\> Get-OrchestratorCheckOutStatus -Runbook '{bc9bcb31-8999-4a59-a080-c6142337a4d5}'

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$RunbookGUID
    )

    Process {
        foreach ($runbook in $RunbookGuid) {
            try {
                if (!$script:connectionHandle) {
                    Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
                }

                If ($PSCmdlet.ShouldProcess("")) {
                    $oStatus = New-Object object
                    $status = New-Object Runtime.InteropServices.VariantWrapper($oStatus)
                    #GetCheckOutStatus                           void GetCheckOutStatus (int, string, Variant)
                    write-verbose ("Getting CheckoutStatus for runbook {0} using connectionHandle {1}" -f $runbook, $script:connectionHandle)
                    $oismgr.GetCheckOutStatus($script:connectionHandle, $runbook, [ref]$Status)
                    $status
                }
            }
            catch {
                Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
            }
        }
    }
}

Function Set-OrchestratorCheckIn {
    <#
        .SYNOPSIS
            Set-OrchestratorCheckIn

        .DESCRIPTION
            Returns whether a policy is running or not

        .PARAMETER TransactionID
            The transaction guid

        .PARAMETER ObjectId
            The object guid

        .PARAMETER Comment
            The check in comment

        .EXAMPLE
            PS C:\> Set-OrchestratorCheckIn -TransactionId ([guid]::NewGuid()) -ObjectId '{bc9bcb31-8999-4a59-a080-c6142337a4d5}' -Comment "Testing DW"

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TransactionId,

        [Parameter(Position=1, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ObjectID,

        [Parameter(Position=2, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Comment
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                #void CheckIn (Handle, bstrTransactionID, bstrObjectID, bstrComment)
                $oismgr.CheckIn($script:connectionHandle, $TransactionId, $ObjectID, $Comment)
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Set-OrchestratorCheckOut {
    <#
        .SYNOPSIS
            Set-OrchestratorCheckOut

        .DESCRIPTION
            Returns whether a policy is running or not

        .PARAMETER  ObjectId
            The object id

        .PARAMETER  Options
            The options field

        .EXAMPLE
            PS C:\> Set-OrchestratorCheckOut -ObjectId '{bc9bcb31-8999-4a59-a080-c6142337a4d5}'

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ObjectID,

        [Parameter(Position=1, Mandatory=$false, ValueFromPipeLine=$true)]
        [System.String]$Options
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oObjectData = New-Object object
                $ObjectData = New-Object Runtime.InteropServices.VariantWrapper($oObjectData)

                #CheckOut (lHandle As Long, bstrObjectID As String, bstrOptions As String, pvarObjectData)
                $oismgr.CheckOut($script:connectionHandle, $ObjectID, $Options, [ref]$ObjectData)

                $ObjectData
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function UndoOrchestratorCheckOut {
    <#
        .SYNOPSIS
            UndoOrchestratorCheckOut

        .DESCRIPTION
            Un-does a check out to an orchestrator runbook

        .PARAMETER  ObjectId
            The object id

        .PARAMETER  Options
            The options field

        .EXAMPLE
            PS C:\> UndoOrchestratorCheckOut -ObjectId '{bc9bcb31-8999-4a59-a080-c6142337a4d5}'

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ObjectID,

        [Parameter(Position=1, Mandatory=$false, ValueFromPipeLine=$true)]
        [System.String]$Options
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oObjectData = New-Object object
                $ObjectData = New-Object Runtime.InteropServices.VariantWrapper($oObjectData)

                #UndoCheckOut (lHandle As Long, bstrObjectID As String, lOptions As Long, pvarObjectData)
                $oismgr.UndoCheckOut($script:connectionHandle, $ObjectID, $Options, [ref]$ObjectData)

                $ObjectData
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorActionServerTypes {
    <#
        .SYNOPSIS
            Get-OrchestratorActionServerTypes

        .DESCRIPTION
            Returns Orchestrator Server Action Types

        .EXAMPLE
            PS C:\> Get-OrchestratorActionServerTypes

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oActionServerTypes = New-Object object
                $ActionServerTypes = New-Object Runtime.InteropServices.VariantWrapper($oActionServerTypes)

                #void GetActionServerTypes (int, pvarActionServerTypes)
                $oismgr.GetActionServerTypes($script:connectionHandle, [ref]$ActionServerTypes)

                $ActionServerTypes
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorActionServers {
    <#
        .SYNOPSIS
            Get-OrchestratorActionServers

        .DESCRIPTION
            Returns Orchestrator Server Action Types

        .PARAMETER  Type
            The Type

        .EXAMPLE
            PS C:\> Get-OrchestratorActionServers

        .EXAMPLE
            PS C:\> Get-OrchestratorActionServers -Type "UnknownValue"

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, ValueFromPipeLine=$true)]
        [System.String]$Type
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oActionServers = New-Object object
                $ActionServers = New-Object Runtime.InteropServices.VariantWrapper($oActionServers)

                #void GetActionServerTypes (int, pvarActionServerTypes)
                $oismgr.GetActionServers($script:connectionHandle, $Type, [ref]$ActionServers)

                $ActionServers
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorFolderContents {
    <#
        .SYNOPSIS
            Get-OrchestratorFolderContents

        .DESCRIPTION
            Gets the contents of an orchestrator folder

        .PARAMETER FolderGuid
            The Folder id

        .EXAMPLE
            PS C:\> Get-OrchestratorFolderContents -FolderGuid '{ba3393e8-17bb-428a-840b-2612d92296b1}'

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FolderGuid
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oFolderContents = New-Object object
                $FolderContents = New-Object Runtime.InteropServices.VariantWrapper($oFolderContents)

                #void GetFolderContents (lHandle As Long, bstrFolderID As String, pvarFolderContents)
                $oismgr.GetFolderContents($script:connectionHandle, $FolderGuid, [ref]$FolderContents)

                # $FolderContents

                $xmlFolderContents = [XML]$FolderContents
                $xmlFolderContents.OuterXml | Format-XML
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorFolderContentsV2 {
    <#
        .SYNOPSIS
            OrchestratorFolderContentsV2

        .DESCRIPTION
            Gets the contents of an orchestrator folder

        .PARAMETER FolderGuid
            The Folder id

        .PARAMETER Policies
            Get Policies

        .PARAMETER Folders
            Get Folders

        .EXAMPLE
            PS C:\> Get-OrchestratorFolderContents -FolderGuid '{ba3393e8-17bb-428a-840b-2612d92296b1}'

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FolderGuid,

        [Parameter(ParameterSetName="specific")]
        [Switch]$Policies,

        [Parameter(ParameterSetName="specific")]
        [Switch]$Folders
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oFolderContents = New-Object object
                $FolderContents = New-Object Runtime.InteropServices.VariantWrapper($oFolderContents)

                #void GetFolderContents (lHandle As Long, bstrFolderID As String, pvarFolderContents)
                $oismgr.GetFolderContents($script:connectionHandle, $FolderGuid, [ref]$FolderContents)

                $xmlFolderContents = [XML]$FolderContents
                # $xmlFolderContents.OuterXml | Format-XML

                If($psCmdlet.ParameterSetName -eq '__AllParameterSets') {
                    return $xmlFolderContents.OuterXml | Format-XML
                }
                elseif ($psCmdlet.ParameterSetName -eq 'specific') {
                    [XML]$out = $null

                    if ($Folders) {

                    }

                    if ($Policies) {

                    }
                    return $out.OuterXml | Format-XML
                }
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorFolderPathFromID {
    <#
        .SYNOPSIS
            Get-OrchestratorFolderPathFromID

        .DESCRIPTION
            Gets the folder path from a folder id

        .PARAMETER  FolderGuid
            The Folder id

        .EXAMPLE
            PS C:\> Get-OrchestratorFolderPathFromID -FolderGuid '{ba3393e8-17bb-428a-840b-2612d92296b1}'

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FolderGuid
    )

    Process {
        try {
            If ($PSCmdlet.ShouldProcess("")) {
                if ($null -eq $script:oisMgr) { $script:oisMgr = new-object -com OpalisManagementService.OpalisManager}

                $oFolderPath = New-Object object
                $FolderPath = New-Object Runtime.InteropServices.VariantWrapper($oFolderPath)

                #GetFolderPathFromID (bstrFolderID As String, pvarFolderPath)
                $oismgr.GetFolderPathFromID($FolderGuid, [ref]$FolderPath)

                $FolderPath
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorFolderByPath {
    <#
        .SYNOPSIS
            Get-OrchestratorFolderByPath

        .DESCRIPTION
            Gets the folder XML by using the folder path

        .PARAMETER FolderPath
            Path to the folder of which to return the XML representation. Form of Policies\FolderName\SubFolderName

        .EXAMPLE
            PS C:\> Get-OrchestratorFolderByPath -Folderpath '/Policies/SomethingTBD'

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Folderpath
        #TODO Add aliases
    )

    Process {
        try {
            If ($PSCmdlet.ShouldProcess("")) {
                if ($null -eq $script:oisMgr) { $script:oisMgr = new-object -com OpalisManagementService.OpalisManager}

                [XML]$folders = Get-OrchestratorFolders
                if ( $null -eq $folders) { Write-Error "Error in $($MyInvocation.MyCommand): `nNo folders returned." }
                #XMLDOCUMENT

                #ToDo finish
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorFolders {
    <#
        .SYNOPSIS
            Get-OrchestratorFolders

        .DESCRIPTION
            Gets the orchestrator folders

        .EXAMPLE
            PS C:\> Get-OrchestratorFolders

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oFolders = New-Object object
                $Folders = New-Object Runtime.InteropServices.VariantWrapper($oFolders)

                #void GetFolders (lHandle As Long, bstrFolderID As String, pvarFolders)
                $oismgr.GetFolders($script:connectionHandle, $ResourceFolderRoot.Runbooks, [ref]$Folders)
                [XML]$xmlFolders = $Folders

                $xmlFolders | Format-XML
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorSubFolders {
    <#
        .SYNOPSIS
            Get-OrchestratorSubFolders

        .DESCRIPTION
            Gets the orchestrator folders

        .PARAMETER  FolderGuid
            The Folder id

        .EXAMPLE
            PS C:\> Get-OrchestratorSubFolder -FolderGuid '{ba3393e8-17bb-428a-840b-2612d92296b1}'

        .OUTPUTS
            System.String

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FolderGuid
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $oFolders = New-Object object
                $Folders = New-Object Runtime.InteropServices.VariantWrapper($oFolders)

                #void GetFolders (lHandle As Long, bstrFolderID As String, pvarFolders)
                $oismgr.GetFolders($script:connectionHandle, $FolderGuid, [ref]$Folders)

                $Folders
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function New-OrchestratorFolder {
    <#
        .SYNOPSIS
            New-OrchestratorFolder

        .DESCRIPTION
            Creates an orchestrator folder

        .PARAMETER FolderName
            The Folder Name

        .PARAMETER FolderDescription
            The Folder Description

        .PARAMETER ParentFolderId
            The Parent Folder id

        .PARAMETER FolderGuid
            The FolderGuid

        .EXAMPLE
            PS C:\> New-OrchestratorFolder -FolderName "David123" -FolderDescription "DavidsTestFolder" -ParentFolderId 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .EXAMPLE
            PS C:\> New-OrchestratorFolder -FolderName "David123" -FolderDescription "DavidsTestFolder" -ParentFolderId 'ba3393e8-17bb-428a-840b-2612d92296b1' -FolderGuid 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("ParentFolderGuid")]
        [GUID]$ParentFolderId,

        [Parameter(Position=1, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FolderName,

        [Parameter(Position=2, ValueFromPipeLine=$true)]
        [GUID]$FolderGuid = [guid]::NewGuid(),

        [Parameter(Position=3, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Description")]
        [System.String]$FolderDescription
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($FolderName)) {

                # Use variant wrapper on inbound XML
                [XML]$addFolderXML = "<Folder>" +
                    "<UniqueID datatype=`"string`"></UniqueID>" +
                    "<Name datatype=`"string`"></Name>" +
                    "<Description datatype=`"string`"></Description>" +
                    "<ParentID datatype=`"string`"></ParentID>" +
                    "<TimeCreated datatype=`"date`"></TimeCreated>" +
                    "<CreatedBy datatype=`"string`"></CreatedBy>" +
                    "<LastModified datatype=`"date`"></LastModified>" +
                    "<LastModifiedBy datatype=`"null`"></LastModifiedBy>" +
                    "<Disabled datatype=`"bool`">FALSE</Disabled>" +
                    "</Folder>"

                $addFolderXML.Folder.Name.InnerText = $folderName
                $addFolderXML.Folder.Description.InnerText = $folderDescription
                $addFolderXML.Folder.UniqueID.InnerText = $folderGuid.ToString("B")

                $FolderData = New-Object Runtime.InteropServices.VariantWrapper($addFolderXML.InnerXml)

                # void AddFolder (lHandle As Long, bstrParentID As String, pvarFolderData)
                $oismgr.AddFolder($script:connectionHandle, $ParentFolderId.ToString("B"), [ref]$FolderData)

                [XML]$xmlFolderData = $folderData
                $xmlFolderData
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function New-OrchestratorResource {
    <#
        .SYNOPSIS
            New-OrchestratorResource

        .DESCRIPTION
            Creates an orchestrator resource

        .PARAMETER ParentId
            The Parent Id

        .PARAMETER ResourceData
            The resource data xml

        .EXAMPLE
            PS C:\> New-OrchestratorResource -ParentId 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("ParentGuid")]
        [GUID]$ParentId,

        [Parameter(Position=1, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Data")]
        [XML]$ResourceData
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {
                $ResData = New-Object Runtime.InteropServices.VariantWrapper($ResourceData.InnerXml)

                #TODO Get ParentID from passed XML if not explicitly set by using $ResourceData.ParentID.InnerText

                # void AddResource(ByVal lHandle As Long, ByVal bstrParentID As String, ResourceData)
                $oismgr.AddResource($script:connectionHandle, $ParentId.ToString("B"), [ref]$ResData)

            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function ModifyOrchestratorFolder {
    <#
        .SYNOPSIS
            ModifyOrchestratorFolder

        .DESCRIPTION
            Modifies an orchestrator folder

        .PARAMETER FolderName
            The Folder Name

        .PARAMETER FolderDescription
            The Folder Description

        .PARAMETER ParentFolderId
            The Parent Folder id

        .PARAMETER FolderGuid
            The FolderGuid

        .EXAMPLE
            PS C:\> ModifyOrchestratorFolder -FolderName "David123" -FolderDescription "DavidsTestFolder" -ParentFolderId 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .EXAMPLE
            PS C:\> ModifyOrchestratorFolder-FolderName "David123" -FolderDescription "DavidsTestFolder" -ParentFolderId 'ba3393e8-17bb-428a-840b-2612d92296b1' -FolderGuid 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [GUID]$ParentFolderId,

        [Parameter(Position=1, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FolderName,

        [Parameter(Position=2, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FolderDescription,

        [Parameter(Position=3, ValueFromPipeLine=$true)]
        [GUID]$FolderGuid = [guid]::NewGuid()
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess("")) {

                # Use variant wrapper on inbound XML
                [XML]$addFolderXML = "<Folder>" +
                    "<UniqueID datatype=`"string`"></UniqueID>" +
                    "<Name datatype=`"string`"></Name>" +
                    "<Description datatype=`"string`"></Description>" +
                    "<ParentID datatype=`"string`"></ParentID>" +
                    "<TimeCreated datatype=`"date`"></TimeCreated>" +
                    "<CreatedBy datatype=`"string`"></CreatedBy>" +
                    "<LastModified datatype=`"date`"></LastModified>" +
                    "<LastModifiedBy datatype=`"null`"></LastModifiedBy>" +
                    "<Disabled datatype=`"bool`">FALSE</Disabled>" +
                    "</Folder>"

                $addFolderXML.Folder.Name.InnerText = $folderName
                $addFolderXML.Folder.Description.InnerText = $folderDescription
                $addFolderXML.Folder.UniqueID.InnerText = $folderGuid.ToString("B")

                $FolderData = New-Object Runtime.InteropServices.VariantWrapper($addFolderXML.InnerXml)

                # void ModifyFolder (int, bstrFolderID, varFolderData)
                $oismgr.ModifyFolder($script:connectionHandle, $ParentFolderId.ToString("B"), [ref]$FolderData)

                [XML]$xmlFolderData = $folderData
                $xmlFolderData
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Import-OrchestratorPolicyFolders {
    <#
        .SYNOPSIS
            Import-OrchestratorPolicyFolders

        .DESCRIPTION
            Imports policy folder(s)

        .PARAMETER File
            The File Name

        .EXAMPLE
            PS C:\> Import-OrchestratorPolicyFolders -File c:\myfile.ois_export

        .OUTPUTS
            Boolean

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([Boolean])]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [ValidateScript({
            try {
                Get-Item $_ -ErrorAction Stop
            } catch [System.Management.Automation.ItemNotFoundException] {
                Throw [System.Management.Automation.ItemNotFoundException] "${_}"
            }
        })]
        [String]$File
    )

    Process {

            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($File)) {

                [xml]$importXML = Get-Content -Path $File

                # Get all policy folders using an XPath Query thus excluding GlobalSettings
                $FolderNodes = $importXML.SelectNodes("/ExportData/Policies//Folder");
                $i = 1
                foreach ($FolderNode in $FolderNodes)
                {
                    Write-Progress -activity "Importing Policy Folders" -status "Folder: $i of $($FolderNodes.Count)" -percentComplete (($i / $FolderNodes.Count)  * 100)
                    try {
                        $ParentNodeID = $FolderNode.ParentNode.UniqueID
                        if (! $ParentNodeID) { Continue } # Ignore this as it is the root runbook folder thus no parentNode present

                        if (Test-OrchestratorFolderExistence -FolderGuid $FolderNode.UniqueID) {
                            Write-Warning "Warning occurred in $($MyInvocation.MyCommand): Folder Name: $($FolderNode.Name) Unique ID: $($FolderNode.UniqueID) - Exists, Skipping."
                        }
                        else {
                            New-OrchestratorFolder `
                                -FolderName $FolderNode.Name `
                                -ParentFolderId $ParentNodeID `
                                -Description "OIP Import via OrchestratorCOM" `
                                -FolderGuid $FolderNode.UniqueID`
                                -verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) `
                                | out-null
                        }
                    }
                    catch {
                        Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
                    }
                    $i++
                }
            }
        # }
        # catch {
        #     Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        # }
    }
}

Function Import-OrchestratorRunbooks {
    <#
        .SYNOPSIS
            Import-OrchestratorRunbooks

        .DESCRIPTION
            Imports runbooks / policies from an exported runbook file.

        .PARAMETER File
            The File Name

        .EXAMPLE
            PS C:\> Import-OrchestratorRunbooks -File c:\something.ois_export

        .OUTPUTS
            Boolean

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([Boolean])]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [ValidateScript({
            try {
                Get-Item $_ -ErrorAction Stop
            } catch [System.Management.Automation.ItemNotFoundException] {
                Throw [System.Management.Automation.ItemNotFoundException] "${_}"
            }
        })]
        [String]$File
    )

    Begin {
        [xml]$importXML = Get-Content -Path $File
    }

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($file)) {

                # Get all Policies using an XPath Query
                $PolicyNodes = $importXML.SelectNodes("/ExportData/Policies//Policy");

                # Get object types once to avoid subsequent multiple calls (think of these as the available integration pack activities including the inbuilt ones)
                $ObjectTypes = Get-OrchestratorObjectTypes
                $i = 1
                foreach ($PolicyNode in $PolicyNodes)
                {
                    Write-Progress -Id 1 -activity "Importing Runbooks" -status "Runbook: $i of $($PolicyNodes.Count)" -percentComplete (($i / $PolicyNodes.Count)  * 100)
                    If (Test-OrchestratorPolicyExistence -PolicyGuid $PolicyNode.UniqueID.InnerText) {
                        Write-Warning "Warning occurred in $($MyInvocation.MyCommand): Policy Name: $($PolicyNode.Name.InnerText), Unique ID: $($PolicyNode.UniqueID.InnerText) - Exists, Skipping."
                    }
                    else
                    {
                        [xml]$policyToAdd = $PolicyNode.OuterXML

                        # Get the object nodes (think of these as the icons and connectors)
                        $ObjectNodes = $policyToAdd.SelectNodes("//Object")
                        $o = 1
                        foreach ($ObjectNode in $ObjectNodes) {
                            Write-Progress -ParentId 1 -activity "Creating policy objects" -status "Object: $o of $($ObjectNodes.Count)" -percentComplete (($o / $ObjectNodes.Count)  * 100)
                            # Check the object type is available first, otherwise the method with throw a foreign key constraint error when it trys to insert
                            # the row into the DB, previously got object types once, rather than hammer the COM api, this should only change if someone imported
                            # an integration pack in the middle of an import, might be a better way of doing this..

                            if ($ObjectTypes.Guid -NotContains [guid]$objectNode.ObjectType.InnerText ) {
                                Write-Error "Error occurred in $($MyInvocation.MyCommand): ObjectTypes doesn't contain object type: $($objectNode.ObjectType.InnerText)`nIntegration pack missing?" #-ErrorAction Stop
                                continue
                            }else{
                                # Create the objects within the policy
                                New-OrchestratorResource `
                                    -ParentId $ObjectNode.ParentID.InnerText`
                                    -ResourceData $ObjectNode.OuterXML `
                                    -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
                            }
                            $o++
                        }

                        # Create the actual policy / runbook
                        New-OrchestratorPolicy `
                            -ParentFolderId $policyToAdd.Policy.ParentID.InnerText `
                            -PolicyData $policyToAdd `
                            -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
                    }
                    $i++
                }
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Import-OrchestratorGlobalSettingsFolders {
    <#
        .SYNOPSIS
            Import-OrchestratorGlobalSettingsFolders

        .DESCRIPTION
            Imports global settings from an exported runbook file.

        .PARAMETER File
            The File Name

        .EXAMPLE
            PS C:\> Import-OrchestratorGlobalSettingsFolders -File c:\something.ois_export

        .OUTPUTS
            Boolean

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([Boolean])]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [ValidateScript({
            try {
                Get-Item $_ -ErrorAction Stop
            } catch [System.Management.Automation.ItemNotFoundException] {
                Throw [System.Management.Automation.ItemNotFoundException] "${_}"
            }
        })]
        [String]$File
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($file)) {

                [xml]$importXML = Get-Content -Path $File
                $i = 1

                # Create Folders
                $GlobalFolderNodes = $ImportXML.SelectNodes("/ExportData/GlobalSettings//Folder");
                foreach ($GlobalFolderNode in $GlobalFolderNodes)
                {
                    try {
                        Write-Progress -activity "Creating Global Settings Folders" -status "Folder: $i of $($GlobalFolderNodes.Count)" -percentComplete (($i / $GlobalFolderNodes.Count)  * 100)
                        $GlobalParentNodeID = $GlobalFolderNode.ParentNode.UniqueID # Parent Folder Node ID

                        if (! $GlobalParentNodeID) { Continue} # This is root folder thus no parentNode present so skip it.

                        If (Test-OrchestratorFolderExistence -FolderGuid $GlobalFolderNode.UniqueID) {
                            Write-Warning "Folder Name: $($GlobalFolderNode.Name) Unique ID: $($GlobalFolderNode.UniqueID) - Exists, Skipping."
                        }
                        else {
                            New-OrchestratorFolder `
                                -FolderName $GlobalFolderNode.Name `
                                -ParentFolderId $GlobalParentNodeID `
                                -Description "OIP Import via OrchestratorCOM" `
                                -FolderGuid $GlobalFolderNode.UniqueID `
                                -verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) `
                                | out-null
                        }
                    }
                    catch {
                        Write-Error "Exception occurred in $($MyInvocation.MyCommand): Error creating folder `n$($_.Exception)"
                    }
                    $i++
                }
            }
            return $True
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Import-OrchestratorGlobalConfiguration {
    <#
        .SYNOPSIS
           Import-OrchestratorGlobalConfiguration

        .DESCRIPTION
            Imports global configuration data from export file.

        .PARAMETER File
            The File Name

        .EXAMPLE
            PS C:\> Import-OrchestratorGlobalConfiguration -File c:\myfile.ois_export

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [ValidateScript({
            try {
                Get-Item $_ -ErrorAction Stop
            } catch [System.Management.Automation.ItemNotFoundException] {
                Throw [System.Management.Automation.ItemNotFoundException] "${_}"
            }
        })]
        [String]$File
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($File)) {

                [xml]$importXML = Get-Content -Path $File

                $entryNodes = $importXML.SelectNodes("/ExportData/GlobalConfigurations/Entry");
                $i = 1
                foreach ($entryNode in $entryNodes) {
                    Write-Progress -activity "Importing Global Configuration" -status "Entry: $i of $($entryNodes.Count)" -percentComplete (($i / $entryNodes.Count)  * 100)
                    #[XML]$configXML = [System.Web.HttpUtility]::HtmlDecode($entryNode.data)
                    #$configXML | Format-XML
                    if (! (Set-OrchestratorConfigurationValue -ConfigurationData $entryNode.OuterXML -verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) )) {
                         Write-Error "$($MyInvocation.MyCommand): Error setting configuration value"
                    }
                    $i++
                }
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function New-OrchestratorPolicy {
    <#
        .SYNOPSIS
            New-OrchestratorPolicy

        .DESCRIPTION
            Creates an orchestrator policy

        .PARAMETER ParentFolderId
            The Parent Folder id

        .PARAMETER PolicyData
            The XML policy data in the form <POLICY>....</POLICY>

        .EXAMPLE
            PS C:\> New-OrchestratorPolicy --ParentFolderId 'ba3393e8-17bb-428a-840b-2612d92296b1' -PolicyData $ImportedData.PolicyNode[1]

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("ParentFolderGuid")]
        [GUID]$ParentFolderId,

        [Parameter(Position=1,  Mandatory=$true,ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [XML]$PolicyData
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($PolicyData.Policy.Name.InnerText)) {

                $varPolicyData = New-Object Runtime.InteropServices.VariantWrapper($PolicyData.InnerXml)

                # AddPolicy void AddPolicy (int, bstrParentID, pvarPolicyData)
                $oismgr.AddPolicy($script:connectionHandle, $ParentFolderId.ToString("B"), [ref]$varPolicyData)

                #[XML]$xmlPolicyData = $varPolicyData
                #$xmlPolicyData
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function ModifyOrchestratorPolicy {
    <#
        .SYNOPSIS
            ModifyOrchestratorPolicy

        .DESCRIPTION
            Modifies an existing orchestrator policy

        .PARAMETER PolicyGuid
            The policy Guid

        .PARAMETER PolicyData
            The XML policy data in the form <POLICY>....</POLICY>

        .EXAMPLE
            PS C:\> ModifyOrchestratorPolicy --ParentFolderId 'ba3393e8-17bb-428a-840b-2612d92296b1' -PolicyData $ImportedData.PolicyNode[1]

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid","PolicyId")]
        [GUID]$PolicyGuid,

        [Parameter(Position=1,  Mandatory=$true,ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [XML]$PolicyData
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($PolicyData.Policy.Name.InnerText)) {

                #$varPolicyData = New-Object Runtime.InteropServices.VariantWrapper($PolicyData.InnerXml)

                $oUniqueId = New-Object object
                $uniqueId = New-Object Runtime.InteropServices.VariantWrapper($oUniqueId)

                # void ModifyPolicy (lHandle, bstrPolicyID, varPolicyData, pvarUniqueKey)
                $oismgr.ModifyPolicy($script:connectionHandle, $PolicyGuid.ToString("B"), $PolicyData.InnerXml, [ref]$uniqueId)

                # [XML]$xmlPolicyData = $varPolicyData
                # $xmlPolicyData
                $uniqueId
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function ModifyOrchestratorObject {
    <#
        .SYNOPSIS
            ModifyOrchestratorObject

        .DESCRIPTION
            Modifies an orchestrator object

        .PARAMETER ObjectId
            The Object Guid

        .PARAMETER UniqueKey
            The unique key for the object?? TODO: Confirm what this is!!

        .PARAMETER ObjectData
            The XML object data in the form <OBJECT>....</OBJECT>

        .EXAMPLE
            PS C:\> ModifyOrchestratorObject -ObjectId 'ba3393e8-17bb-428a-840b-2612d92296b1' -UniqueKey 'abcdef12-3456-7890-840b-1232d91233c4'-PolicyData $ImportedData.PolicyNode[1]

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid","ObjectGuid")]
        [GUID]$ObjectId,

        [Parameter(Position=2, Mandatory=$true,ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [GUID]$UniqueKey,

        [Parameter(Position=3, Mandatory=$true,ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Object")]
        [XML]$ObjectData
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($PolicyData.Policy.Name.InnerText)) {

                $oObject = New-Object object
                $object = New-Object Runtime.InteropServices.VariantWrapper($oObject)
                # void ModifyObject (lHandle As Long, bstrObjectID As String, bstrUniqueKey As String, varObjectData)
                $oismgr.ModifyObject($script:connectionHandle, $ObjectId.ToString("B"), $UniqueKey.ToString("B"), $PolicyData.InnerXml, [ref]$object)

                $uniqueId
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorPolicyPublishState {
    <#
        .SYNOPSIS
            Get-OrchestratorPolicyPublishState

        .DESCRIPTION
            Gets the folder XML by using the folder path

        .PARAMETER PolicyGuid
            The PolicyGuid

        .EXAMPLE
            PS C:\> Get-OrchestratorPolicyPublishState -PolicyGuid 47BDA5DC-15FD-42B9-9AE9-70B54E22A1F0

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid")]
        [GUID[]]$PolicyGuid
    )

    Process {
        foreach ($pol in $PolicyGuid) {
            try {
                if (!$script:connectionHandle) {
                    Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
                }

                If ($PSCmdlet.ShouldProcess($pol)) {
                    $oFlags = New-Object object
                    $Flags = New-Object Runtime.InteropServices.VariantWrapper($oFlags)

                    # GetPolicyPublishState void GetPolicyPublishState (lHandle As Long, ByVal bstrPolicyID As String, plFlags As Long)
                    $oismgr.GetPolicyPublishState($script:connectionHandle, $pol.ToString("B"), [ref]$flags)
                    $flags
                }
            }
            catch {
                Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
            }
        }
    }
}

Function Get-OrchestratorPolicy {
    <#
        .SYNOPSIS
            Get-OrchestratorPolicy

        .DESCRIPTION
            Gets the policy XML

        .PARAMETER PolicyGuid
            The PolicyGuid

        .EXAMPLE
            PS C:\> Get-OrchestratorPolicy -PolicyGuid 47BDA5DC-15FD-42B9-9AE9-70B54E22A1F0

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid")]
        [GUID[]]$PolicyGuid
    )

    Process {
        foreach ($pol in $PolicyGuid) {
            try {
                if (!$script:connectionHandle) {
                    Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
                }

                If ($PSCmdlet.ShouldProcess($pol)) {
                    $oPolicyData = New-Object object
                    $policyData = New-Object Runtime.InteropServices.VariantWrapper($oPolicyData)

                    #  LoadPolicyvoid LoadPolicy (int, bstrPolicyID, pvarPolicyData)
                    $oismgr.LoadPolicy($script:connectionHandle, $pol.ToString("B"), [ref]$policyData)
                    $policyData
                }
            }
            catch {
                Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
            }
        }
    }
}

Function Get-OrchestratorObjectTypes {
    <#
        .SYNOPSIS
            Get-OrchestratorObjectTypes

        .DESCRIPTION
            Gets the object types

        .EXAMPLE
            PS C:\> Get-OrchestratorObjectTypes

        .OUTPUTS
            PSCustomObject

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([PSCustomObject])]
    param(
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            If ($PSCmdlet.ShouldProcess($pol)) {
                $ObjectTypesArray = @() # array to store the results

                $oObjectDetails = New-Object object
                $objectDetails = New-Object Runtime.InteropServices.VariantWrapper($oObjectDetails)

                # void GetObjectTypes (handle, pvarObjectDetails) # Gets list of Object Types from datastore
                $oismgr.GetObjectTypes($script:connectionHandle, [ref]$objectDetails)

                #$objectTypes = [XML]$objectDetails
                $objectTypeNodes = ([XML]$objectDetails).SelectNodes("//ObjectType")

                foreach ($objectTypeNode in $objectTypeNodes) {
                    $ObjectTypesArray += [PSCustomObject]@{
                            'Name'=[string]$objectTypeNode.Name.InnerText;
                            'Description'=[string]$objectTypeNode.Description.InnerText;
                            'Guid'=[Guid]$objectTypeNode.UniqueID.InnerText
                    }
                }

                return $ObjectTypesArray

            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Get-OrchestratorConfigurationValue {
    <#
        .SYNOPSIS
            Get-OrchestratorConfigurationValue

        .DESCRIPTION
            Gets configuration values

        .PARAMETER ConfigurationId
            The configuration Id to retrieve

        .PARAMETER NoFormatting
            Switch parameter as to whether to disable formatting of the the xml output

        .EXAMPLE
            PS C:\> Get-OrchestratorConfigurationValue -ConfigrationId '143587E7-595B-499D-A13E-00E9BD02F059'

        .EXAMPLE
            PS C:\> Get-OrchestratorConfigurationValue -ConfigrationId '{143587E7-595B-499D-A13E-00E9BD02F059}'

        .EXAMPLE
            PS C:\> Get-OrchestratorConfigurationValue '{143587E7-595B-499D-A13E-00E9BD02F059}'

        .EXAMPLE
            PS C:\> Get-OrchestratorConfigurationValue 143587E7-595B-499D-A13E-00E9BD02F059

        .OUTPUTS
            XML

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Scope="Function")]
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([XML])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid","ConfigurationGuid")]
        [GUID[]]$ConfigurationId,

        [Switch]$NoFormatting
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            foreach ($configId in $ConfigurationId) {
                If ($PSCmdlet.ShouldProcess($configId)) {
                    #$ObjectTypesArray = @()

                    $oValues = New-Object object
                    $values = New-Object Runtime.InteropServices.VariantWrapper($oValues)

                    # GetConfigurationValues(ByVal lHandle As Long, ByVal bstrConfigID As String, pvarValues)
                    $oismgr.GetConfigurationValues($script:connectionHandle, $configId.ToString("B"), [ref]$values)

                    <#
                    $objectTypes = [XML]$objectDetails
                    $configurationNodes = ([XML]$objectDetails).SelectNodes("//ObjectType")

                    foreach ($objectTypeNode in $objectTypeNodes) {
                        $ObjectTypesArray += [PSCustomObject]@{
                                'Name'=[string]$objectTypeNode.Name.InnerText;
                                'Description'=[string]$objectTypeNode.Description.InnerText;
                                'Guid'=[Guid]$objectTypeNode.UniqueID.InnerText
                        }
                    }

                    return $ObjectTypesArray
                    #>

                    If ($NoFormatting) {
                        return $values
                    }else {
                        return ($values | Format-XML)
                    }
                }
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function Set-OrchestratorConfigurationValue {
    <#
        .SYNOPSIS
            Set-OrchestratorConfigurationValue

        .DESCRIPTION
            Set configuration values

        .PARAMETER ConfigurationData
            XML Configuration Data in the form <ENTRY></ENTRY>

        .EXAMPLE
            PS C:\> Set-OrchestratorConfigurationValue $ConfigXML

        .EXAMPLE
            PS C:\> Set-OrchestratorConfigurationValue -ConfigurationData $ConfigXML

        .EXAMPLE
            PS C:\> Set-OrchestratorConfigurationValue -ConfigurationData $entryNode.OuterXML

        .OUTPUTS
            Boolean

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([Boolean])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [XML[]]$ConfigurationData
    )

    Process {
        try {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }

            foreach ($config in $ConfigurationData) {
                $configId = [Guid]$config.SelectSingleNode("/Entry/ID").InnerText
                If ($PSCmdlet.ShouldProcess($configId)) {

                    #Un-escape the contained XML - might not be needed, but leave here for the moment.
                    #[xml]$data = [System.Web.HttpUtility]::HtmlDecode($config.Entry.Data)

                    # $configValues = New-Object Runtime.InteropServices.VariantWrapper($config)
                    # void SetConfigurationValues (lHandle As Long, bstrConfigID As String, varValues
                    $oismgr.SetConfigurationValues($script:connectionHandle, $configId.ToString("B"), $config.OuterXML)
                    return $true
                }
            }
        }
        catch {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }
    }
}

Function isOrchestratorFolderDeleted {
    <#
        .SYNOPSIS
            isOrchestratorFolderDeleted

        .DESCRIPTION
            Checks the database to see if the folder is in a deleted state

        .PARAMETER FolderGuid
            The FolderGuid

        .PARAMETER DatabaseName
             The orchestrator database name

        .PARAMETER DatabaseServer
             The orchestrator database server name

        .EXAMPLE
            PS C:\> isOrchestratorFolderDeleted -FolderGuid 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .OUTPUTS
            System.Boolean

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid")]
        [GUID[]]$FolderGuid,

        [Parameter(Position=1, ValueFromPipeLine=$true)]
        [Alias("Server")]
        [String]$DatabaseServer = "localhost",

        [Parameter(Position=2, ValueFromPipeLine=$true)]
        [Alias("Database")]
        [String]$DatabaseName = "Orchestrator"
    )

    Begin {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }
    }

    Process {
        foreach ($folder in $FolderGuid) {
            try {
                If ($PSCmdlet.ShouldProcess($folder)) {

                    $connStringBuilder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
                    $connStringBuilder["Data Source"] = $DatabaseServer
                    $connStringBuilder["Initial Catalog"] = $DatabaseName
                    $connStringBuilder["Integrated Security"] = $true

                    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStringBuilder.ConnectionString

                    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand "select [UniqueID],[Name],[Description],[Disabled],[Deleted],[ParentID] FROM [FOLDERS] WHERE [UniqueID] = '$Folder'", $SqlConnection
                    $SqlCmd.CommandTimeout = 0

                    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd

                    $DataSet = New-Object System.Data.DataSet
                    $RowCount = $SqlAdapter.Fill($DataSet)
                    $SqlConnection.Close()

                    If ($RowCount -lt 1) { Return ("Folder {0} not found in the database" -f $folder.ToString("B")) }
                    If ($RowCount -gt 1) { Write-Error "$($MyInvocation.MyCommand): Too much data returned, aborting" -ErrorAction Stop }

                    # Found entry

                    foreach ($Row in $dataset.Tables[0].Rows)
                    {
                        If ($Row.Deleted -eq "True") {
                             Write-verbose ("Folder {0} is showing as deleted in the database" -f $folder.ToString("B"))
                             # Todo this doesn't feel correct!
                             return $true
                        }
                    }
                }
            }
            catch {
                throw
            }
        }
    }
    End {
        $SqlConnection.Dispose()
    }
}

Function ExportOrchestratorGlobalConfigurationToSQLScript {
    <#
        .SYNOPSIS
            ExportOrchestratorConfigurations

        .DESCRIPTION
            Exports the configuration settings from the Orchestrator Database

        .PARAMETER DatabaseName
             The orchestrator database name

        .PARAMETER DatabaseServer
             The orchestrator database server name

        .EXAMPLE
            PS C:\> ExportOrchestratorConfigurations

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    param(
        # [Parameter(Position=0, Mandatory=$True,ValueFromPipeline=$true)]
        # [ValidateScript({
        #     try {
        #         Get-Item $_ -ErrorAction Stop
        #     } catch [System.Management.Automation.ItemNotFoundException] {
        #         Throw [System.Management.Automation.ItemNotFoundException] "${_}"
        #     }
        # })]
        # [String]$File,

        [Parameter(Position=0, ValueFromPipeLine=$true)]
        [Alias("Server")]
        [String]$DatabaseServer = "localhost",

        [Parameter(Position=1, ValueFromPipeLine=$true)]
        [Alias("Database")]
        [String]$DatabaseName = "Orchestrator"
    )

    Begin {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }
    }

    Process {
        try {
            If ($PSCmdlet.ShouldProcess($DatabaseServer)) {

                $connStringBuilder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
                $connStringBuilder["Data Source"] = $DatabaseServer
                $connStringBuilder["Initial Catalog"] = $DatabaseName
                $connStringBuilder["Integrated Security"] = $true

                $SqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStringBuilder.ConnectionString

                $SqlCmd = New-Object System.Data.SqlClient.SqlCommand "select * FROM [CONFIGURATION] WHERE [DataName] = 'Configurations'", $SqlConnection
                $SqlCmd.CommandTimeout = 0

                $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd

                $DataSet = New-Object System.Data.DataSet
                $RowCount = $SqlAdapter.Fill($DataSet)
                $SqlConnection.Close()

                If ($RowCount -lt 1) { Write-Error ("No configurations found in the database") -ErrorAction Stop }

                ##Write-Output "USE [$DatabaseName]`nGO"
                foreach ($Row in $dataset.Tables[0].Rows)
                {
                    Write-Output ("INSERT [dbo].[CONFIGURATION] ([UniqueID], [TypeGUID], [DataName], [DataValue], [IsPassword]) VALUES (N'{0}', N'{1}', N'{2}', N'{3}', {4})" -f $Row.UniqueID, $Row.TypeGUID, $Row.DataName, $Row.DataValue, $Row.IsPassword.GetHashCode())
                }
            }
        }
        catch {
            throw
        }

    }
    End {
        $SqlConnection.Dispose()
    }
}

Function ExecuteSQLScript {
    <#
        .SYNOPSIS
            ExecuteSQLScript

        .DESCRIPTION
            Executes a sql script against the orchestrator database

        .PARAMETER File
             The script file name

        .PARAMETER DatabaseName
             The orchestrator database name

        .PARAMETER DatabaseServer
             The orchestrator database server name

        .EXAMPLE
            PS C:\> ExecuteSQLScript -File Configurations.sql

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    param(
        [Parameter(Position=0, Mandatory=$True,ValueFromPipeline=$true)]
        [ValidateScript({
            try {
                Get-Item $_ -ErrorAction Stop
            } catch [System.Management.Automation.ItemNotFoundException] {
                Throw [System.Management.Automation.ItemNotFoundException] "${_}"
            }
        })]
        [String]$File,

        [Parameter(Position=1, ValueFromPipeLine=$true)]
        [Alias("Server")]
        [String]$DatabaseServer = "localhost",

        [Parameter(Position=2, ValueFromPipeLine=$true)]
        [Alias("Database")]
        [String]$DatabaseName = "Orchestrator"
    )

    Begin {
    }

    Process {
        try {
            If ($PSCmdlet.ShouldProcess($File)) {

                $SQL = Get-Content $File

                $connStringBuilder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
                $connStringBuilder["Data Source"] = $DatabaseServer
                $connStringBuilder["Initial Catalog"] = $DatabaseName
                $connStringBuilder["Integrated Security"] = $true

                $SqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStringBuilder.ConnectionString

                $SqlCmd = New-Object System.Data.SqlClient.SqlCommand $SQL, $SqlConnection
                $SqlCmd.CommandTimeout = 0

                $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd

                $DataSet = New-Object System.Data.DataSet
                $RowCount = $SqlAdapter.Fill($DataSet)
                $SqlConnection.Close()

                foreach ($Row in $dataset.Tables[0].Rows)
                {
                    $Row
                }
            }
        }
        catch {
            throw
        }

    }
    End {
        $SqlConnection.Dispose()
    }
}

Function Test-OrchestratorFolderExistence {
    <#
        .SYNOPSIS
            Test-OrchestratorFolderExistence

        .DESCRIPTION
            Checks the database to see if the folder exists

        .PARAMETER FolderGuid
            The FolderGuid

        .PARAMETER DatabaseName
             The orchestrator database name

        .PARAMETER DatabaseServer
             The orchestrator database server name

        .EXAMPLE
            PS C:\> Test-OrchestratorFolderExistence -FolderGuid 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .OUTPUTS
            System.Boolean
        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid")]
        [GUID[]]$FolderGuid,

        [Parameter(Position=1, ValueFromPipeLine=$true)]
        [Alias("Server")]
        [String]$DatabaseServer = "localhost",

        [Parameter(Position=2, ValueFromPipeLine=$true)]
        [Alias("Database")]
        [String]$DatabaseName = "Orchestrator"
    )

    Begin {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }
    }

    Process {
        foreach ($folder in $FolderGuid) {
            try {
                If ($PSCmdlet.ShouldProcess($folder)) {

                    $connStringBuilder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
                    $connStringBuilder["Data Source"] = $DatabaseServer
                    $connStringBuilder["Initial Catalog"] = $DatabaseName
                    $connStringBuilder["Integrated Security"] = $true

                    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStringBuilder.ConnectionString

                    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand "select [UniqueID],[Name],[Description],[Disabled],[Deleted],[ParentID] FROM [FOLDERS] WHERE [UniqueID] = '$Folder'", $SqlConnection
                    $SqlCmd.CommandTimeout = 0

                    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd

                    $DataSet = New-Object System.Data.DataSet
                    $RowCount = $SqlAdapter.Fill($DataSet)
                    $SqlConnection.Close()

                    If ($RowCount -eq 0) { Return $false }
                    elseif ($RowCount -eq 1) { Return $true }

                    # Shouldn't get this far, but just in case.
                    Throw "Error"
                }
            }
            catch {
                throw
            }
        }
    }
    End {
        $SqlConnection.Dispose()
    }
}

Function Test-OrchestratorPolicyExistence {
    <#
        .SYNOPSIS
            Test-OrchestratorPolicyExistence

        .DESCRIPTION
            Checks to see if a policy exists

        .PARAMETER PolicyGuid
            The PolicyGuid

        .PARAMETER DatabaseName
             The orchestrator database name

        .PARAMETER DatabaseServer
             The orchestrator database server name

        .EXAMPLE
            PS C:\> Test-OrchestratorPolicyExistence -PolicyGuid 'ba3393e8-17bb-428a-840b-2612d92296b1'

        .EXAMPLE
            PS C:\> Test-OrchestratorPolicyExistence -PolicyGuid 'ba3393e8-17bb-428a-840b-2612d92296b1' -DatabaseName "Orchestrator"

        .OUTPUTS
            System.Boolean

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
           https://github.com/davidwallis3101/OrchestratorCOM
    #>

    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Guid")]
        [GUID[]]$PolicyGuid,

        [Parameter(Position=1, ValueFromPipeLine=$true)]
        [Alias("Server")]
        [String]$DatabaseServer = "localhost",

        [Parameter(Position=2, ValueFromPipeLine=$true)]
        [Alias("Database")]
        [String]$DatabaseName = "Orchestrator"
    )

    Begin {
            if (!$script:connectionHandle) {
                Write-Error "$($MyInvocation.MyCommand): Not Connected" -ErrorAction Stop
            }
    }

    Process {
        foreach ($policy in $PolicyGuid) {
            try {
                If ($PSCmdlet.ShouldProcess($policy)) {

                    # $oPol = New-Object object
                    # $pol = New-Object Runtime.InteropServices.VariantWrapper($oPol)

                    #DoesPolicyExist void DoesPolicyExist (bstrPolicyID)
                    #$oismgr.DoesPolicyExist($policy.ToString("B")) # Doesn't seem to work.. cant get data back, resort to SQL..

                    $connStringBuilder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
                    $connStringBuilder["Data Source"] = $DatabaseServer
                    $connStringBuilder["Initial Catalog"] = $DatabaseName
                    $connStringBuilder["Integrated Security"] = $true

                    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStringBuilder.ConnectionString

                    # Get more data than needed as this may then be used for a get-OrchestratorPolicyZZZ function in the future.
                    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand "SELECT [UniqueID],[CreationTime], [LastModified], [LastModifiedBy], [CheckOutTime], [CheckOutLocation], [Deleted] FROM [$DatabaseName].[dbo].[POLICIES] WHERE UniqueID = '$($Policy.ToString("D"))'", $SqlConnection
                    $SqlCmd.CommandTimeout = 0

                    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd

                    $DataSet = New-Object System.Data.DataSet
                    $RowCount = $SqlAdapter.Fill($DataSet)
                    $SqlConnection.Close()

                    If ($RowCount -eq 0) { Return $false }
                    elseif ($RowCount -eq 1) { return $true }
                }
            }
            catch {
                throw
            }
        }
    }
    End {
        $SqlConnection.Dispose()
    }
}

Function Install-OrchestratorIntegrationPack {
    <#
    .SYNOPSIS

        Install-OrchestratorIntegrationPack

    .DESCRIPTION

        Registers and optionally deploys an Integration Pack to the current computer.
        Assumes the current computer is a Management Server (and a Designer/Runbook Server
        in the case of deploying the IP)

        Found here: https://blogs.technet.microsoft.com/orchestrator/2012/05/24/more-fun-with-com-importing-integration-packs-via-powershell/

    .PARAMETER Filename
        The path and filename of the OIP file to be imported

    .PARAMETER Deploy
        Switch parameter used when you want to also deploy the IP.

    .EXAMPLE
        Install-OrchestratorIntegrationPack -OIPFile "C:\Files\Test.OIP" -Deploy

    .OUTPUTS

    #>
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true)]
        [String] $Filename,

        [Parameter(Mandatory=$false)]
        [Switch] $Deploy
    )
    BEGIN
    {
        $sh = new–object -com shell.application
    }
    PROCESS
    {
        try
        {
            if (!$script:connectionHandle) {
                Throw "Not Connected"
            }

            if ($(Test-Path $filename) -eq $false)
            {
                Write-Error "File Not Found!"
                return
            }
            [System.IO.FileInfo]$oipFile = get-item $filename

            $extractDir = $(Join-Path -Path $oipFile.DirectoryName -ChildPath $oipFile.BaseName)
            if (Test-Path $extractDir)
            {
               remove-item $extractDir -Recurse -Force
            }

            $ipSourceDirObj = $sh.namespace($oipFile.DirectoryName)
            $ipSourceDirObj.NewFolder($oipFile.BaseName)
            $extractDirObj = $sh.namespace($extractDir)

            Write-Debug "`n`nExtracting files from the OIP"
            $zipFileName = Join-Path $oipFile.DirectoryName "$($oipFile.BaseName).zip"
            if (Test-Path $zipFileName)
            {
               remove-item $zipFileName -Force
            }
            $zipFile = $oipFile.CopyTo($zipFileName)
            $zipFileObj = $sh.namespace($zipFile.FullName)

            $extractDirObj.CopyHere($zipFileObj.Items(),8 -and 16 -and 256)

            $commonFiles = ${Env:CommonProgramFiles(x86)}
            if ($null -eq $commonFiles)
            {
                $commonFiles = ${Env:CommonProgramFiles}
            }
            $PacksDir = Join-Path $commonFiles "Microsoft System Center 2012\Orchestrator\Management Server\Components\Packs"
            $ObjectsDir = Join-Path $commonFiles "Microsoft System Center 2012\Orchestrator\Management Server\Components\Objects"

            if ($(Test-Path $PacksDir) -eq $false)
            {
                Write-Error "Could not find $($PacksDir)"
                return
            }
            if ($(Test-Path $ObjectsDir) -eq $false)
            {
                Write-Error "Could not find $($ObjectsDir)"
                return
            }

            # Copy the MSI File to %Common Files%\Microsoft System Center 2012\Orchestrator\Management Server\Components\Objects
            $msiFile = $extractDirObj.self | Get-ChildItem | Where-Object {$_.extension -eq ".msi"}

            $newMSIfile = $(Join-Path $ObjectsDir $msiFile.Name)
            if (Test-Path $newMSIfile)
            {
               remove-item $newMSIfile -Force
            }

            $msiFile.CopyTo($newMSIfile)
            $productName = Get-MSIProperty -Filename $msiFile.FullName -Propertyname "ProductName"
            $productCode =  Get-MSIProperty -Filename $msiFile.FullName -Propertyname "ProductCode"

            #now use the MgmtService to install the IP
            [System.IO.FileInfo]$capfile = $extractDirObj.self | Get-ChildItem | Where-Object {$_.extension -eq ".cap"}
            if ($capfile)
            {
                Write-Verbose "Extracting $($capFile)"
                $capXml = New-Object XML
                $capXml.Load($capfile.Fullname)

                # Need to modify the CAP file to add Product ID and Product Name because Deployment Manager
                # reads the MSI file for this and inserts it into the DB so it displays in the UI. The COM
                # interface does not do this, so it needs to be done manually if you want it displayed.
                #
                #     <ProductName datatype="string">IP_SYSTEMCENTERDATAPROTECTIONMANAGER_1.0.OIP</ProductName>
                #    <ProductID datatype="string">{9422FCC6-11C4-4827-AC49-C5FD352C8AA0}</ProductID>

                [Xml]$prodName = "<ProductName datatype=`"string`">$($ProductName)</ProductName>"
                [Xml]$prodID = "<ProductID datatype=`"string`">$($ProductCode)</ProductID>"

                $c = $capXml.ImportNode($prodName.ProductName, $true)
                $d = $capXml.ImportNode($prodID.ProductID, $true)

                $capXml.Cap.AppendChild($c)
                $capXml.Cap.AppendChild($d)

                #$oIPinfo = new–object object
                [ref]$ipinfo = New-Object Runtime.InteropServices.VariantWrapper($capXml.get_innerxml())
                Write-Verbose "Importing Integration Pack $($capFile)"
                #$retval = $oismgr.AddIntegrationPack($script:connectionHandle, $ipinfo)
                $oismgr.AddIntegrationPack($script:connectionHandle, $ipinfo) | out-null
            }

            # Copy the OIP File to the %Common Files%\Microsoft System Center 2012\Orchestrator\Management Server\Components\Packs
            # directory and change the name to the GUID of the IP
            $productCodeOipFilename = "$($ProductCode).OIP"

            $newOIPfile = $(Join-Path $PacksDir $productCodeOipFilename)
            if (Test-Path $newOIPfile)
            {
               remove-item $newOIPfile -Force
            }
            $oipFile.CopyTo($newOIPfile)

            if ($PSBoundParameters.ContainsKey('Deploy') -eq $false)
            {
                return
            }

            if ($(Test-Path $newMSIfile) -eq $false)
            {
                Write-Error "Could not find $($newMSIfile)"
                return
            }

            Write-Verbose "Running msiexec to install $($newMSIfile)"

            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo.FileName = "msiexec.exe"
            $proc.StartInfo.Arguments =  "/i `"$($newMSIfile)`" /qn"
            $proc.Start() | out–null
            $proc.WaitForExit()

        }
        catch
        {
            Write-Error "Exception occurred in $($MyInvocation.MyCommand): `n$($_.Exception)"
        }

    }

}


Function Format-XML{
    <#
        .SYNOPSIS
            Format-XML

        .DESCRIPTION
            Formats XML to be more friendly on the eyes

        .PARAMETER XML
            The Input XML

        .PARAMETER INDENT
            The Indentation level

        .EXAMPLE
            PS C:\> get-content c:\test.xml | Format-XML

        .INPUTS
            XML

        .LINK
            https://blogs.msdn.microsoft.com/powershell/2008/01/18/format-xml/

        .NOTES
            DW - Modified to keep attributes on the same line for ease of reading.
            DW - Modified to omit the xml declaration
    #>
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [xml]$xml,

        [Parameter(Position=1)]
        [int]$indent=2
    )

    # Create a settings object so that we can set NewLIneOnAttributes
    $XMLSettings = New-Object System.Xml.XmlWriterSettings
    $XMLSettings.Indent = $true
    $XMLSettings.NewLineOnAttributes = $false
    $XMLSettings.OmitXmlDeclaration = $true

    $StringWriter = New-Object System.IO.StringWriter
    $XMLWriter = [System.Xml.XmlTextWriter]::Create($StringWriter, $XMLSettings)
    $xml.WriteContentTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()

    Write-Output $StringWriter.ToString()
}

Function Get-MsiProperty {
    <#
        .SYNOPSIS
            Get-MsiProperty

        .DESCRIPTION
            Gets MSI Properties

        .PARAMETER FileName
            The MSI Filename

        .PARAMETER PropertyName
             The property name to get

        .EXAMPLE
            PS C:\> Get-MsiProperty -Filename 'c:\test.msi' -PropertyName

        .EXAMPLE
            PS C:\> Get-MsiProperty-Filename $msiFile.FullName -Propertyname "ProductName"

        .EXAMPLE
            PS C:\> Get-MsiProperty -Filename $msiFile.FullName -Propertyname "ProductCode"

        .NOTES
            For additonal information please see https://github.com/davidwallis3101/OrchestratorCOM

        .LINK
    #>

    Param (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        [String]$Filename,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        [String]$Propertyname
    )

    # A quick check to see if the file exists
    if(!(Test-Path $Filename))
    {
        throw "Could not find " + $Filename
    }

    # Create an empty hashtable to store properties in
    $msiProp = ""
    # Creating WI object and load MSI database
    $wiObject = New-Object -com WindowsInstaller.Installer
    $wiDatabase = $wiObject.InvokeMethod("OpenDatabase", (Resolve-Path $Filename).Path, 0)
    # Open the Property-view
    $view = $wiDatabase.InvokeMethod("OpenView", "SELECT * FROM Property")
    $view.InvokeMethod("Execute")
    # Loop thru the table
    $r = $view.InvokeMethod("Fetch")
    while($null -ne $r)
    {
        # Add property and value to hash table
        $prop = $r.InvokeParamProperty("StringData",1)
        $value = $r.InvokeParamProperty("StringData",2)
        if ($prop -eq $PropertyName)
        {
            $msiProp = $value
        }
        # Fetch the next row
        $r = $view.InvokeMethod("Fetch")
    }
    $view.InvokeMethod("Close")

    return $msiProp
}
###################################################################################################################

<# Likely needed

    GetPolicyIDFromPath                     void GetPolicyIDFromPath (bstrPolicyPath, pvarPolicyID)
    GetPolicyPathFromID                     void GetPolicyPathFromID (bstrPolicyID, bstrPolicyID)
    GetPolicyObjectList                     void GetPolicyObjectList (int, bstrPolicyID, pvarObjectList)

    GetPolicyRunningState                   void GetPolicyRunningState (varSeqNumber, pvarPolicyRunning)
    GetPolicyRunStatus                      void GetPolicyRunStatus (int, bstrExecInstanceID, pvarPolicyDetails)

    AddPolicy                               void AddPolicy (int, bstrParentID, pvarPolicyData)
    ModifyPolicy                            void ModifyPolicy (int, bstrPolicyID, varPolicyData, pvarUniqueKey)
    DeletePolicy                            void DeletePolicy (int, string, int)
    LoadPolicy                              void LoadPolicy (int, bstrPolicyID, pvarPolicyData)

    GetAuditHistory                         void GetAuditHistory (int, bstrObjectID, bstrTransactionID, bstrCommand, pvarData)
    GetLogHistory                           void GetLogHistory (int, bstrPolicyID, lFlags, pvarLogData)
    GetLogHistoryObjectDetails              void GetLogHistoryObjectDetails (int, bstrInstanceID, bstrObjectID, bstrInstanceNumber, pvarLogData)
    GetLogHistoryObjects                    void GetLogHistoryObjects (int, bstrPolicyID, bstrInstanceID, pvarLogData)
    GetLogObjectDetails                     void GetLogObjectDetails (int, bstrObjectID, bstrInstanceID, bstrInstanceNumber, pvarObjectInformation)
    GetEventDetails                         void GetEventDetails (bstrUniqueID, pvarEvents)

    AddIntegrationPack                      void AddIntegrationPack (int, Variant)
    RemoveIntegrationPack                   void RemoveIntegrationPack (int, Variant)
#>

# Export Module Functions
Export-ModuleMember -Function Connect-OrchestratorComInterface
Export-ModuleMember -Function Disconnect-OrchestratorComInterface
Export-ModuleMember -Function Get-OrchestratorIntegrationPacks
Export-ModuleMember -Function Get-OrchestratorPoliciesWithoutImages
Export-ModuleMember -Function Get-OrchestratorPolicyRunningState
Export-ModuleMember -Function Get-OrchestratorCheckOutStatus
Export-ModuleMember -Function Set-OrchestratorCheckIn
Export-ModuleMember -Function Set-OrchestratorCheckOut
Export-ModuleMember -Function UndoOrchestratorCheckOut
Export-ModuleMember -Function Get-OrchestratorActionServerTypes
Export-ModuleMember -Function Get-OrchestratorActionServers
Export-ModuleMember -Function Get-OrchestratorFolderContents
Export-ModuleMember -Function Get-OrchestratorFolderContentsV2
Export-ModuleMember -Function Get-OrchestratorFolderPathFromID
Export-ModuleMember -Function Get-OrchestratorFolderByPath
Export-ModuleMember -Function Get-OrchestratorFolders
Export-ModuleMember -Function Get-OrchestratorSubFolders
Export-ModuleMember -Function New-OrchestratorFolder
Export-ModuleMember -Function New-OrchestratorResource
Export-ModuleMember -Function ModifyOrchestratorFolder
Export-ModuleMember -Function Import-OrchestratorPolicyFolders
Export-ModuleMember -Function Import-OrchestratorRunbooks
Export-ModuleMember -Function Import-OrchestratorGlobalSettingsFolders
Export-ModuleMember -Function Import-OrchestratorGlobalConfiguration
Export-ModuleMember -Function New-OrchestratorPolicy
Export-ModuleMember -Function ModifyOrchestratorPolicy
Export-ModuleMember -Function ModifyOrchestratorObject
Export-ModuleMember -Function Get-OrchestratorPolicyPublishState
Export-ModuleMember -Function Get-OrchestratorPolicy
Export-ModuleMember -Function Get-OrchestratorObjectTypes
Export-ModuleMember -Function Get-OrchestratorConfigurationValue
Export-ModuleMember -Function Set-OrchestratorConfigurationValue
Export-ModuleMember -Function isOrchestratorFolderDeleted
Export-ModuleMember -Function ExportOrchestratorGlobalConfigurationToSQLScript
Export-ModuleMember -Function Test-OrchestratorFolderExistence
Export-ModuleMember -Function Test-OrchestratorPolicyExistence
Export-ModuleMember -Function ExecuteSQLScript
