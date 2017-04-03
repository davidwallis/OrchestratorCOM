# OrchestratorCOM

This is a powershell module for wrapping the System Center Orchestrator (formally Opalis) COM object.

Sources/References:
* https://blogs.technet.microsoft.com/orchestrator/2012/04/30/getting-deeper-information-from-orchestrator-via-com/
* https://blogs.technet.microsoft.com/orchestrator/2012/05/01/getting-deeper-information-from-orchestrator-via-com-part-2/
* https://blogs.technet.microsoft.com/orchestrator/2012/05/01/getting-deeper-information-from-orchestrator-via-com-part-3/
* https://blogs.technet.microsoft.com/orchestrator/2012/05/29/ips-and-activities-and-how-those-are-represented-in-the-database/
* https://blogs.technet.microsoft.com/orchestrator/2012/05/24/more-fun-with-com-importing-integration-packs-via-powershell/
* https://scorch.codeplex.com/SourceControl/latest

Requires:

* Powershell V4 (5.x issues see below)
* 32 Bit powershell process.

I haven't yet written an export function as I am currently manually doing that part - or using stuff that is available here: https://scorch.codeplex.com/SourceControl/latest
I suspect this will change in the coming weeks as I plan to use a SQL trigger to monitor the last modified date on policy objects

Warning:

I am not currently bothering to change the GUID's - why bother if exporting from one server to another (assuming your arent sharing a DB instance!)

Exporting of global configurations doesnt seem to currently work as I expected so I extract the configurations from the database and generate a sql insert statement, bear in mind that passwords will not transfer across due to the use of AES encryption keys in orchestrator,
I suspect you could back up and restore the Symetric and Asysmetric keys that are used, these can be located by looking at the relevant stored procedures in the orchestrator database or by running:

```sql
SELECT * FROM sys.asymmetric_keys WHERE name = 'ORCHESTRATOR_ASYM_KEY'
SELECT * FROM sys.symmetric_keys WHERE name = 'ORCHESTRATOR_SYM_KEY'
```

# Importing a ois_export file

## Warnings
* This will not be supported by Myself or Microsoft so backup your system and Database first
* Ensure that the required Integration packs are already installed, this module contains functions for doing this kind of work but are currently not fully tested.
* It is assumed that the destination is a blank Orchestrator install with just the integration packs present, it will not currently update existing runbooks (I'm currently testing by reverting snapshots of the server)


## Usage
Launch an x86 Powershell prompt on your target orchestrator server
Import the module assuming it's in your PSModulePath
```powershell
import-module OrchestratorCOM
```

Connect to the COM api
```powershell
Connect-OrchestratorComInterface -Credential (Get-credential domain\user)
```

Import Policy Folders, Global Settings Folders, Policies and Global Configuration
```powershell
$oisFile = 'c:\everything.ois_export'

Import-OrchestratorPolicyFolders -File $oisFile
Import-OrchestratorRunbooks -File $oisFile
Import-OrchestratorGlobalSettingsFolders -File $oisFile
# TODO Write Import for global settings
Import-OrchestratorGlobalConfiguration -File $oisFile
```

As a one of task the global configurations can be exported from the source server by running the following command on the SOURCE server
```powershell
ExportOrchestratorGlobalConfigurationToSQLScript | out-file c:\Configurations.SQL
```

These can then be imported using the following command on the Destination Server, you would then need to ammend any credentials due to encryption errors, if anyone can shed any light on how to deal with this I'm all ears as it must be possible using the Salt available within the OIS file.

```powershell
#Import Configuration Crude!!
ExecuteSQLScript c:\Configurations.SQL
```

Finally disconnect from the COM api.

```powershell
Disconnect-OrchestratorComInterface -verbose
```

I am currently testing this so feel free to help add functionality if requrired or suggest changes as I know it's not perfect as yet.

NB: I had issues getting the variant wrapper to work in WMF5, - I have logged this on user voice here: https://windowsserver.uservoice.com/forums/301869-powershell/suggestions/17787103-wmf-5-invalid-callee-when-using-variant-wrapper
