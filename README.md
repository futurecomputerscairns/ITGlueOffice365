# ITGlueOffice365
Huge shoutout to GCITS, this script is largely based on their work over at https://github.com/GCITS/knowledge-base/tree/master/ITGlue/Office365Sync

This is a project to import Office 365 information for each customer you have global admin credentials for in ITGlue.
The 'password type' for the global admin credentials must be set to 'Microsoft Office 365 Admin'

## Prerequisites 

Powershell 3.0 or greater is required for this script to run.

**Due to requirements I have in my environments, I needed to manually download the module and place into C:\Temp\ITGlue\Modules**

Edit line #45 to 'Import-Module MSOnline' if you don't face this issue.

## Running the script
In order to get started, please follow [this guide](https://github.com/GCITS/knowledge-base/blob/master/ITGlue/Office365Sync/README.md#how-to-sync-office-365-tenant-info-with-it-glue) to create the Office 365 Flexible Asset in ITGlue.


Once this is created, the script can be executed as follows:
```
.\Office365ITGlue.ps1 -Key "ITG.[YOURAPIKEYHERE]"
