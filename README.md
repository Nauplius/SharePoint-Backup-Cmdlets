**Repository Moved**

This repository has moved to [tseward/SharePoint-Backup-Cmdlets](https://github.com/tseward/SharePoint-Backup-Cmdlets). 
 
 Installation
* Extract the zip file on the SharePoint server
* From the SharePoint Management Shell, run SharePointBAC.ps1
* Once solution deployment is complete, close the SharePoint Management Shell and re-open it prior to usage

 Usage

SharePointBAC provides the following cmdlets:
* Get-SPBackupCatalog
* Set-SPBackupCatalog
* Remove-SPBackupCatalog
* Export-SPBackupCatalog

This is an example of leveraging SharePointBAC to automatically trim the backups on disk to 2 days and email backup status to an Administrator and Backup Operator:


    Add-PSSnapin Microsoft.SharePoint.PowerShell
    $cat = Get-SPBackupCatalog \\backupserver\SharePointBackup
    Backup-SPFarm -Directory \\backupserver\SharePointBackup -BackupMethod Full
    $cat.Refresh()
    $cat | Remove-SPBackupCatalog -RetainCount 2 -Confirm:$false
    $cat.Refresh()
    $cat | Send-SPBackupStatus -Recipients "admin@example.com,backupop@example.com"


 Uninstallation
* Navigate to Central Administration as a Farm Administrator
* Under System -> Manage farm solutions, click on sharepointbac.wsp.  Click on Retract.  Once retracted, click on sharepointbac.wsp again and click Remove
