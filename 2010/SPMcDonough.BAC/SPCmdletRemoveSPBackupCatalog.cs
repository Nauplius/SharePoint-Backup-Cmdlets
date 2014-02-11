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


    [Cmdlet(VerbsCommon.Remove, "SPBackupCatalog", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    [SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true, RequireUserMachineAdmin = false)]
    [CmdletDescription("This cmdlet is used to selectively remove backup sets from a backup catalog - a process known as \"grooming.\" This permits retention of backup sets based on a count or size limit, and it prevents a backup catalog from growing over time to consume all storage available to it.")]
    [Example(Code = "Remove-SPBackupCatalog \\\\SPFarmShare\\Backups -RetainCount 3", Remarks = "In this example, all backups in the catalog located at \\\\SPFarmShare\\Backup are deleted save for the last three non error-containing full-mode backups (-RetainCount 3) and any differential backups that depend on them. Any restores that were performed remain in the backup catalog; only backup sets/folders are deleted.")]
    [Example(Code = "Get-SPBackupCatalog \\\\SPFarmShare\\Backups | Remove-SPBackupCatalog -RetainSize 10GB", Remarks = "In this example, an SPBackupCatalog object is created and piped to the Remove-SPBackupCatalog cmdlet with a -RetainSize of 10GB specified. The cmdlet then examines the backup catalog; if it finds that the catalog is larger than 10GB in size, it will delete the oldest full backup set, any differential backups that depend on the full backup, and any configuration-only backups that took place prior to the next-oldest full-mode backup. The cmdlet then re-examines the size of the backup catalog and repeats the process until the backup catalog is equal to or less than 10GB in size. Note that by default, only non error-containing backup sets are considered during size calculations to avoid a situation where all viable backup sets are groomed-out. To avoid this behavior, use the -IgnoreBackupSetErrors switch.")]
    [Example(Code = "Remove-SPBackupCatalog \\\\SPFarmShare\\Backups -RetainCount 5 -RetainSize 20GB", Remarks = "This example combines both a -RetainCount and -RetainSize to groom backups to a fixed full backup count and size. The backup catalog is first trimmed down to five full-mode non error-containing backups and any differential backups that depend on those full-mode backups. After being groomed to five full-mode backups, the backup catalog is further trimmed (if necessary) to get its size to 20GB or less based on remaining non error-containing backup sets.")]
    [Example(Code = "Remove-SPBackupCatalog \\\\SPFarmShare\\Backups -EntireCatalog", Remarks = "In this example, all backup data is deleted from the catalog that is targeted. An spbrtoc.xml file remains in the catalog location, and it continues to reflect restore and non-backup operations, but all backup set data in the location is deleted. This frees the maximum amount of disk space. Note that if this switch is specified, the -RetainCount and -RetainSize parameters have no effect.")]
    [Example(Code = "Remove-SPBackupCatalog \\\\SPFarmShare\\Backups -RetainCount 1 -IgnoreBackupSetErrors", Remarks = "In this example, all backup sets save for the most recent one are deleted. The -IgnoreBackupSetErrors means that any backup set errors will be ignored during the grooming process. Even if the most recent backup set contains errors, it will be the only backup set that is retained following execution of this command. This illustrates why use of the -IngoreBackupSetErrors switch should be considered carefully; it could leave you without a viable backup from which to restore.")]
    [RelatedCmdlets(typeof(SPCmdletGetSPBackupCatalog))]
    public class SPCmdletRemoveSPBackupCatalog : SPRemoveCmdletBase<SPBackupCatalog>
    {


        #region Overrides (SPRemoveCmdletBase)


        /// <summary>
        /// This method is called when the cmdlet is invoked, and it's where we take care of examining
        /// the supplied backup catalog (as an <see cref="SPBackupCatalog"/>) to determine which backup
        /// sets, if any, should be removed from it.
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        protected override void DeleteDataObject()
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
                // The path we take depends on whether or not the entire catalog should be deleted.
                if (this.EntireCatalog)
                {
                    // See if we need to confirm processing.
                    if (isProcessingConfirmed == null)
                    {
                        isProcessingConfirmed = ShouldProcess(workingCatalog.CatalogPath, res.SPCmdletRemoveSPBackupCatalog_ShouldProcess1);
                    }
                    // Carry out delete if we've gotten the thumbs-up
                    if (isProcessingConfirmed.HasValue && isProcessingConfirmed == true)
                    {
                        // Empty out the backup catalog; once it completes, we'll need to refresh the
                        // working catalog due to changes the method makes.
                        BackupCatalogUtilities.EmptyBackupCatalog(workingCatalog);
                        workingCatalog.Refresh();
                    }
                }
                else
                {
                    // Something less than the entire catalog is to be deleted. Ensure that we've got 
                    // some usable combination of parameters.
                    if (this.RetainCount == 0 && this.RetainSize == 0)
                    {
                        psError.Category = ErrorCategory.InvalidArgument;
                        psError.Object = null;
                        throw new ArgumentException(res.SPCmdletRemoveSPBackupCatalog_Ex_Arg1);
                    }

                    // First pass at retention is by count if specified; of course, we only need
                    // to process if the number of full backups is greater than the retain number.
                    if (this.RetainCount > 0 && workingCatalog.FullBackupCount > this.RetainCount)
                    {
                        // See if we need to confirm processing.
                        if (isProcessingConfirmed == null)
                        {
                            isProcessingConfirmed = ShouldProcess(workingCatalog.CatalogPath, res.SPCmdletRemoveSPBackupCatalog_ShouldProcess1);
                        }
                        // Carry out delete if we've gotten the thumbs-up
                        if (isProcessingConfirmed.HasValue && isProcessingConfirmed == true)
                        {
                            // Carry out the trim operation; once it completes, we'll need to refresh the
                            // working catalog due to changes the method makes.
                            BackupCatalogUtilities.TrimBackupCatalogCount(workingCatalog, this.RetainCount, this.IgnoreBackupSetErrors);
                            workingCatalog.Refresh();
                        }
                    }

                    // Next pass at retention is based on the desired size of all backups combined.
                    if (this.RetainSize > 0 && workingCatalog.CatalogSize > this.RetainSize)
                    {
                        // See if we need to confirm processing.
                        if (isProcessingConfirmed == null)
                        {
                            isProcessingConfirmed = ShouldProcess(workingCatalog.CatalogPath, res.SPCmdletRemoveSPBackupCatalog_ShouldProcess1);
                        }
                        // Carry out delete if we've gotten the thumbs-up
                        if (isProcessingConfirmed.HasValue && isProcessingConfirmed == true)
                        {
                            BackupCatalogUtilities.TrimBackupCatalogSize(workingCatalog, this.RetainSize, this.IgnoreBackupSetErrors);
                            workingCatalog.Refresh();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Something went wrong with the backup catalog, the actual grooming, or something in-between.
                // Regardless, we need to terminate the pipeline -- nothing more can be done.
                ThrowTerminatingError(ex, psError.Category, psError.Object);
            }
        }


        #endregion Overrides (SPRemoveCmdletBase)


        #region Properties


        [Parameter(
            HelpMessage = "Specifying this switch results in the complete deletion of all data in the backup catalog. When this switch is specified, the -RetainCount and -RetainSize parameters have no effect.")]
        public SwitchParameter EntireCatalog { get; set; }


        /// <summary>
        /// This property represents the backup catalog that will be used for operations. Since the
        /// input is an <see cref="SPBackupCatalogPipeBind"/>, either a live <see cref="SPBackupCatalog"/>
        /// can be supplied - or some other object that can be coerced/interpreted by the pipe bind type.
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPBackupCatalogPipeBind"/>
        [Parameter(
            HelpMessage = "The backup catalog which will serve as the target for operations. This can be specified as an existing SPBackupCatalog object or as a path to the target backup catalog root (i.e., where the spbrtoc.xml resides).",
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true)]
        [ValidateNotNullOrEmpty]
        public SPBackupCatalogPipeBind Identity { get; set; }


        [Parameter(
            HelpMessage = "When -RetainCount and -RetainSize calculations are performed, they don't include backup sets that have errors. This is to avoid scenarios where backup grooming may remove good backup sets (i.e., those without errors) and leave only bad backup sets behind. If the -IgnoreBackupSetErrors switch is specified, then -RetainCount and -RetainSize calculation don't distinguish between backup sets with errors and those without. Use of this switch should be considered carefully, as it could leave you in a situation where your backup catalog contains only bad backups that are non-restorable.")]
        public SwitchParameter IgnoreBackupSetErrors { get; set; }
        
        
        /// <summary>
        /// This property represents the maximum number of full backups sets that should be retained
        /// following any trim or clean-up operations. A minimum value of 1 must be supplied for this
        /// property; a value of zero would mean that all backups are deleted, and we don't permit that.
        /// </summary>
        [Parameter(
            HelpMessage = "The maximum number of non error-containing full backup sets to retain. When this parameter is specified, only content+configuration full-mode backups are counted for purposes of retention. Configuration-only backups are ignored for purposes of grooming, but they will be groomed-out if they are older than the last full backup to be retained. Differential backups are associated with their full backup for retention purposes. Even though they aren't included in the backup set count, differential backups are groomed-out or retained according to what happens with their full-mode base backup set.",
            Mandatory = false,
            Position = 1,
            ValueFromPipeline = false)]
        [ValidateRange(1,Int32.MaxValue)]
        public Int32 RetainCount { get; set; }


        /// <summary>
        /// This property represents the maximum total size of backup sets that should be retained
        /// following any trim or clean-up operations. A minimum value of 1 must be supplied for this
        /// property, but in reality any values supplied should be much larger since the value is in
        /// bytes.
        /// </summary>
        [Parameter(
            HelpMessage = "The maximum size, in bytes, of non error-containing backup sets to be retained. When this parameter is specified, it will groom a backup catalog to the specified size or below. Since this grooming typically occurs after backups have been performed, it obviously cannot permit backup catalogs from growing to a size greater than the -RetainSize; it simply trims the backup catalogs down after-the-fact.",
            Mandatory = false,
            Position = 2,
            ValueFromPipeline = false)]
        [ValidateRange(1,Int64.MaxValue)]
        public Int64 RetainSize { get; set; }
        

        #endregion Properties


    }
}
