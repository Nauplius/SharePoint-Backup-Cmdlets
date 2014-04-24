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


    [Cmdlet(VerbsCommunications.Send, "SPBackupStatus", ConfirmImpact = ConfirmImpact.Low, SupportsShouldProcess = true)]
    [SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true, RequireUserMachineAdmin = false)]
    [CmdletDescription("This cmdlet is used to send an e-mail to one or more recipients to communicate the backup status of a specific backup catalog. This is commonly done following a backup to let administrators know that a backup has taken place, if the backup contains errors, where the backup set was created, and more. E-mail messages are sent using the farm's configured outbound e-mail settings.")]
    [Example(Code = "Send-SPBackupStatus \\\\SPFarmShare\\Backups -Recipients admin@nosite.com", Remarks = "This example represents the simplest usage scenario. A backup status e-mail for the backup catalog located at \\\\SPFarmShare\\Backups will be sent to admin@nosite.com. The message will contain information about the the most recent backup performed. Details such as the number of errors encountered, the location of the backup set, when the backup started, and more are included.")]
    [Example(Code = "Get-SPBackupCatalog \\\\SPFarmShare\\Backups | Send-SPBackupStatus -Recipients \"admin@nosite.com\" -OnErrorOnly", Remarks = "This example illustrates piping an SPBackupCatalog object to the Send-SPBackupStatus cmdlet to send an e-mail status to admin@nosite.com. Since the -OnErrorOnly switch is employed, and e-mail only gets sent if the last backup set actually had an error count greater than zero.")]
    [Example(Code = "Send-SPBackupStatus E:\\Backups -Recipients \"admin@nosite.com, BackupGroup@nosite.com, FarmAdmin@nosite.com\" -Subject \"Last Backup Status for Production Farm\"", Remarks = "This example demonstrates how an e-mail can be sent with an alternate subject line. It also illustrates how to send the status message to multiple recipients by separating e-mail addresses using commas.")]
    [RelatedCmdlets(typeof(SPCmdletGetSPBackupCatalog))]
    public class SPCmdletSendSPBackupStatus : SPSetCmdletBase<SPBackupCatalog>
    {


        #region Overrides (SPSetCmdletBase)


        /// <summary>
        /// This method is called when the cmdlet is invoked, and it's where we take care of setting up an
        /// e-mail message and sending it out.
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
                // First check to see if this should be an error-only e-mail; if so, we need to ensure that the 
                // we only proceed if an error exists in the latest backup set.
                if (!this.OnErrorOnly || workingCatalog.LastBackupErrorCount > 0)
                {
                    // Validation attributes will ensure that we've got a consistent set of parameters; the only thing
                    // we need to do before proceeding is see if an e-mail should be sent (if appropriate).
                    if (isProcessingConfirmed == null)
                    {
                        isProcessingConfirmed = ShouldProcess(workingCatalog.CatalogPath,res.SPCmdletSendSPBackupStatus_ShouldProcess1);
                    }

                    // Setup the e-mail subject, but override it if an alternate subject line is specified
                    String mailSubject = String.Format(res.SPCmdletSendSPBackupStatus_DefaultSubject, workingCatalog.CatalogPath);
                    if (!String.IsNullOrEmpty(this.Subject))
                    {
                        mailSubject = this.Subject;
                    }

                    // Create the e-mail body that will communicate the status and then send the e-mail.
                    String mailBody = BackupCatalogUtilities.BuildEmailStatusBody(workingCatalog);
                    CommunicationUtilities.SendFarmEmail(this.Recipients, String.Empty, mailSubject, mailBody, true);
                }
            }
            catch (Exception ex) //string not valid datetime
            {
                // Something went wrong with the backup catalog, the e-mail process, or something in-between.
                // Regardless, we need to terminate the pipeline -- nothing more can be done.
                ThrowTerminatingError(ex, psError.Category, psError.Object);
            }

        }


        #endregion Overrides (SPSetCmdletBase)


        #region Properties


        /// <summary>
        /// This property represents the backup catalog that will be the target for status messages. Since the
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
            HelpMessage = "When this switch is specified, a backup status e-mail message is only sent in the event that one or more errors are encountered in the most recent backup performed. Without this switch, a backup status e-mail is always sent.")]
        public SwitchParameter OnErrorOnly { get; set; }


        [Parameter(
            HelpMessage = "The e-mail addresses of the recipients to whom the backup status e-mail will be sent. If multiple e-mail addresses are specified, they must be separated by commas.",
            Mandatory = true,
            Position = 1,
            ValueFromPipeline = false)]
        [ValidateNotNullOrEmpty()]
        public String Recipients { get; set; }


        [Parameter(
            HelpMessage = "The subject line that will be used for the e-mail status message that is sent. If no subject line is specified, a generic subject line will be used.",
            Mandatory = false,
            Position = 2,
            ValueFromPipeline = false)]
        [ValidateNotNullOrEmpty()]
        public String Subject { get; set; }

        
        #endregion Properties


    }
}
