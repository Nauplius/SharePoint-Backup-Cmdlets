#region Namespace Imports


// Standard class imports
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// Imports that are needed for SharePoint and PowerShell
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;

// For documentation generation using Gary Lapointe's MAML generator
using Lapointe.PowerShell.MamlGenerator.Attributes;

// Set resources alias
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{


    [Cmdlet(VerbsCommon.Set, "SPBackupCatalog", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true, RequireUserMachineAdmin = false)]
    [CmdletDescription("This cmdlet provides a mechanism to make changes and updates to specific properties on the spbrtoc.xml file that is associated with a backup catalog.")]
    [Example(Code = "Get-SPBackupCatalog \\\\SPFarmShare\\NewBackupLocation | Set-SPBackupCatalog -UpdateCatalogPath", Remarks = "This example begins by getting the backup catalog that is located at the root of the \\\\SPFarmShare\\NewBackupLocation share. The SPBackupCatalog object that is created is then piped to the Set-SPBackupCatalog command with an option to -UpdateCatalogPath. The use of the -UpdateCatalogPath switch with the Set-SPBackupCatalog cmdlet will result in the replacement of all internal backup set path references to a previous backup catalog path with the current catalog location; i.e., \\\\SPFarmShare\\NewBackupLocation. With this replacement operation, SharePoint can then back up to and restore from the backup catalog (at \\\\SPFarmShare\\NewBackupLocation) as if it had been the original destination for all backup operations.")]
    [RelatedCmdlets(typeof(SPCmdletGetSPBackupCatalog))]
    public class SPCmdletSetSPBackupCatalog : SPSetCmdletBase<SPBackupCatalog>
    {


        #region Overrides (SPSetCmdletBase)


        /// <summary>
        /// This method is called when the cmdlet is invoked, and it's where we take care of parsing
        /// property sets/updates for the backup set and then apply them.
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>        
        protected override void UpdateDataObject()
        {
            // Setup error object in case we need to terminate the pipeline; also create the flag
            // that will be used to support ShouldProcess functionality.
            PowerShellError psError = new PowerShellError();
            Boolean? isProcessingConfirmed = null;

            // Read-in the needed backup catalog object for further operations. Pipebind object
            // will ensure validity prior to return and throw a terminating exception if invalid.
            SPBackupCatalog workingCatalog = this.Identity.Read();
            workingCatalog.Refresh();

            // Until further notice, we assume errors are tied to our catalog and unspecified.
            psError.Object = workingCatalog;
            psError.Category = ErrorCategory.NotSpecified;

            // If anything goes wrong, we need to terminate the pipeline.
            try
            {
                // Validation attributes will ensure that we've got a consistent set of parameters; the only thing
                // we need to do before proceeding is see if an e-mail should be sent (if appropriate).
                if (isProcessingConfirmed == null)
                {
                    isProcessingConfirmed = ShouldProcess(workingCatalog.CatalogPath, res.SPCmdletSetSPBackupCatalog_ShouldProcess1);
                }

                // Get the current catalog path and use that as the basis for modification.
                String catalogPath = workingCatalog.CatalogPath;
                BackupCatalogUtilities.UpdateBackupCatalogPaths(workingCatalog, catalogPath);
            }
            catch (Exception ex)
            {
                // Something went wrong with the backup catalog, the e-mail process, or something in-between.
                // Regardless, we need to terminate the pipeline -- nothing more can be done.
                ThrowTerminatingError(ex, psError.Category, psError.Object);
            }

        }


        #endregion Overrides (SPSetCmdletBase)


        #region Properties


        /// <summary>
        /// This property represents the backup catalog that will be the target for sets/updates. Since the
        /// input is an <see cref="SPBackupCatalogPipeBind"/>, either a live <see cref="SPBackupCatalog"/>
        /// can be supplied - or some other object that can be coerced/interpreted by the pipe bind type.
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPBackupCatalogPipeBind"/>
        [Parameter(
            HelpMessage = "The backup catalog which will serve as the target for property changes. This can be specified as an existing SPBackupCatalog object or as a path to the target backup catalog root (i.e., where the spbrtoc.xml resides).",
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true)]
        [ValidateNotNullOrEmpty]
        public SPBackupCatalogPipeBind Identity { get; set; }


        [Parameter(
            HelpMessage = "When this switch is set, the backup set pointers contained within the catalog will be validated and corrected if necessary to make them consistent with the current path held by the SPBackupCatalog. This type of correction becomes necessary when a backup catalog that was created at one path (for example, \\\\BackupServer\\Backups\\) is moved to another path (for example, \\\\BackupServer\\Restores\\) and referenced. In this scenario, folder pointers that are internal to the backup catalog will still point to \\\\BackupServer\\Backups\\. Use of this switch will correct those pointers to reference the path currently help by the SPBackupCatalog.")]
        public SwitchParameter UpdateCatalogPath { get; set; }


        #endregion Properties


    }
}
