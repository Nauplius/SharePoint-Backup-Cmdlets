#region Namespace Imports


// Standard class imports
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// Imports that are needed for SharePoint and PowerShell
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;

// For documentation generation using Gary Lapointe's MAML generator
using Lapointe.PowerShell.MamlGenerator.Attributes;

// Set resources alias
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{


    /// <summary>
    /// This class represents a backup catalog in SharePoint. "Backup catalog," in this case, is defined
    /// as a location that the catastrophic farm backup API writes to and reads from for purposes of farm
    /// level backups. These locations house an <c>spbrtoc.xml</c> table of contents file and have zero or
    /// more <c>spbrxxxx</c> subdirectories housing individual full and differential backup runs.
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "SPBackupCatalog", ConfirmImpact = ConfirmImpact.Low, SupportsShouldProcess = false)]
    [SPCmdlet(RequireLocalFarmExist=true, RequireUserFarmAdmin=true, RequireUserMachineAdmin=false)]
    [CmdletDescription("This cmdlet is used to retrieve a backup catalog object (SPBackupCatalog) that represents a SharePoint farm backup location. A backup catalog is defined as the top-level folder that contains a table of contents file (spbrtoc.xml) and zero or more backup set folders (as spbrxxxx subfolders). SPBackupCatalog objects provide information about the backup catalog specified and can also be used as input for other operations that modify the backup catalog.")]
    [Example(Code = "Get-SPBackupCatalog", Remarks = "Retrieves a backup catalog for the SharePoint farm's default backup location. The default location is typically specified through the 'Configure Backup Settings' page within the Central Administration site.")]
    [Example(Code = "Get-SPBackupCatalog -CatalogPath \"\\\\SPFarmShare\\Backups\"", Remarks = "In this example, the -CatalogPath parameter is specified to retrieves an SPBackupCatalog for the backup catalog present at \\\\ProdFarm\\Backups UNC path." )]
    [RelatedCmdlets(typeof(SPCmdletExportSPBackupCatalog), typeof(SPCmdletRemoveSPBackupCatalog), typeof(SPCmdletSendSPBackupStatus))]
    public class SPCmdletGetSPBackupCatalog : SPGetCmdletBase<SPBackupCatalog>
    {


        #region Overrides (SPGetCmdletBase)


        protected override IEnumerable<SPBackupCatalog> RetrieveDataObjects()
        {
            // Create a list to return the desired catalog object(s), add the desired catalog,
            // and return it. How we do it depends on whether or not a catalog path was specified.
            List<SPBackupCatalog> catalogList = new List<SPBackupCatalog>();
            SPBackupCatalog newCatalog;
            if (String.IsNullOrEmpty(this.CatalogPath))
            {
                newCatalog = new SPBackupCatalog(SPFarm.Local);
            }
            else
            {
                newCatalog = new SPBackupCatalog(this.CatalogPath);
            }
            catalogList.Add(newCatalog);
            return catalogList;
        }


        #endregion Overrides (SPGetCmdletBase)


        #region Properties


        [Parameter(
            HelpMessage="The file path containing the spbrtoc.xml file for the desired backup catalog. This parameter should normally be specified as a UNC path. In limited circumstances, such as with an all-in-one SharePoint server, a local path (beginning with a drive letter) may be acceptable.",
            Mandatory=false,
            Position=0,
            ValueFromPipeline=false)]
        public String CatalogPath { get; set; }
        
        
        #endregion Properties


    }
}
