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

// Simplifies some I/O and path operations
using System.IO;

// For documentation generation using Gary Lapointe's MAML generator
using Lapointe.PowerShell.MamlGenerator.Attributes;

// Set resources alias
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{

    
    [Cmdlet(VerbsData.Export, "SPBackupCatalog", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true, RequireUserMachineAdmin = false)]
    [CmdletDescription("The Export-SPBackupCatalog cmdlet is commonly used to create an external archive of older backup sets within a backup catalog. Backup sets are included either explicitly through a count or by specifying the number of newer backup sets that should be excluded. Exports that are created contain both the backup sets selected and a table of contents file that is recognized by SharePoint (for later restore operations). Exports are typically created as a single compressed (zip) file, but \"loose exports\" (i.e., without compression) can also be created.")]
    [Example(Code = "Export-SPBackupCatalog \\\\SPFarmShare\\Backups -ExportPath \\\\ArchiveShare\\SharePoint", Remarks = "This example demonstrates how all backup sets in the backup catalog located at the \\\\SPFarmShare\\Backups location can be exported to another share (\\\\ArchiveShare\\SharePoint). The backup sets are compressed into a single zip file which is auto-named (since no -ExportName is specified) using a date/time serial value. Only backup sets that are error-free are exported; if a full backup sets contains errors, any differential backup sets that depends on the full backup set are also omitted.")]
    [Example(Code = "Get-SPBackupCatalog \\\\SPFarmShare\\Backups | Export-SPBackupCatalog -ExportPath \\\\ArchiveShare\\SharePoint -ExportName MyBackupArchive -IncludeCount 1", Remarks = "In this example, an SPBackupCatalog object is created and piped to the Export-SPBackupCatalog cmdlet. The cmdlet then creates an export of the last full backup set (and any differential backup sets tied to it) in the catalog at the specified share: \\\\ArchiveShare\\SharePoint. The export will be found in the share as a single compressed archive file named MyBackupArchive.zip")]
    [Example(Code = "Export-SPBackupCatalog \\\\SPFarmShare\\Backups -ExportPath \\\\ArchiveShare\\SharePoint -ExportName SPExports -ExcludeCount 3 -NoCompression", Remarks = "This command line results in the export of all backup sets in the catalog save for the most recent three (full) sets and any differential backups that depend on them. The -NoCompression switch is applied, so backup sets will not be compressed into a zip archive. Since the -NoCompression switch is used, the -ExportName parameter will be used to create a subdirectory in the -ExportPath share, and exported backup sets will be copied to this location (i.e., \\\\ArchiveShare\\SharePoint\\SPExports).")]
    [Example(Code = "Export-SPBackupCatalog \\\\SPFarmShare\\Backups -ExportPath \\\\ArchiveShare\\SharePoint -IncludeCount 2 -IgnoreBackupSetErrors", Remarks = "This example demonstrates how to export the two oldest full backup sets and any differential backup sets that depend on them. Since no -ExportName is specified, a zip archive filename will automatically be created by the system using date/time information. The -IgnoreBackupSetErrors switch is applied, so backup sets containing errors will not be excluded from the export selection process.")]
    [Example(Code = "Export-SPBackupCatalog \\\\SPFarmShare\\Backups -ExportPath \\\\ArchiveShare\\SharePoint -ExportName NewSPExports -ExcludeCount 2 -NoCompression -UpdateCatalogPath", Remarks = "This command line results in the export of all backup sets in the catalog save for the most recent two (full) sets and any differential backups that depend on them. The -NoCompression switch is applied, so backup sets will not be compressed into a zip archive. Since the -NoCompression switch is used, the -ExportName parameter will be used to create a subdirectory in the -ExportPath share, and exported backup sets will be copied to this location (i.e., \\\\ArchiveShare\\SharePoint\\SPExports). Using the -UpdateCatalogPath switch will adjust all backup set path pointers within the backup catalog to point to the full destination path of the export; i.e., \\\\ArchiveShare\\SharePoint\\SPExports.")]
    [RelatedCmdlets(typeof(SPCmdletGetSPBackupCatalog))]
    public class SPCmdletExportSPBackupCatalog : SPSetCmdletBase<SPBackupCatalog>
    {


        #region Overrides (SPSetCmdletBase)


        /// <summary>
        /// This method is called when the cmdlet is invoked, and it's where we take care of setting up 
        /// and creating the target backup archive/export.
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

            // Some variables we'll need throughout the export process.
            String exportFileOrFolder;
            Int32 exportCount;
            BackupSetExportMode exportMode;

            // If anything goes wrong, we need to terminate the pipeline.
            try
            {
                // Check the export path to make sure that it's someplace we can actually get to.
                if (!Directory.Exists(this.ExportPath))
                {
                    psError.Category = ErrorCategory.InvalidArgument;
                    psError.Object = null;
                    throw new ArgumentException(res.SPCmdletExportSPBackupCatalog_Ex_Arg1);
                }

                // If an export name is specified, make sure that it doesn't contain any invalid
                // path or filename characters; if the export name isn't specified, create one
                // using date/time information.
                if (!String.IsNullOrEmpty(this.ExportName))
                {
                    // If we're using compression, then we need to check against invalid filename
                    // characters; without compression, we're checking against bad path chars.
                    if (this.NoCompression)
                    {
                        Char[] invalidPathChars = Path.GetInvalidPathChars();
                        if (this.ExportName.IndexOfAny(invalidPathChars) != -1)
                        {
                            psError.Category = ErrorCategory.InvalidArgument;
                            psError.Object = null;
                            throw new ArgumentException(res.SPCmdletExportSPBackupCatalog_Ex_Arg2);
                        }
                    }
                    else
                    {
                        Char[] invalidPathChars = Path.GetInvalidFileNameChars();
                        if (this.ExportName.IndexOfAny(invalidPathChars) != -1)
                        {
                            psError.Category = ErrorCategory.InvalidArgument;
                            psError.Object = null;
                            throw new ArgumentException(res.SPCmdletExportSPBackupCatalog_Ex_Arg3);
                        }
                    }

                    // If we're here, then the export name contains only valid characters.
                    exportFileOrFolder = this.ExportName;
                }
                else
                {
                    // No -ExportName was specified, so we'll generate one using the current date/time
                    // and a template.
                    exportFileOrFolder = String.Format(Globals.EXPORT_FILE_OR_PATH_TEMPLATE,
                                                       Globals.BuildFileAndPathCompatibleDateTime(DateTime.Now));
                }

                // Make sure we don't have conflicting -IncludeCount and -ExcludeCount parameter
                // values -- one or the other. Provided we've got acceptable values, we'll parse
                // them out into a usable form for the ultimate export operation.
                if (this.IncludeCount > 0 && this.ExcludeCount > 0)
                {
                    psError.Category = ErrorCategory.InvalidOperation;
                    psError.Object = null;
                    throw new ArgumentException(res.SPCmdletExportSPBackupCatalog_Ex_Arg4);
                }
                else if (this.IncludeCount > 0)
                {
                    exportMode = BackupSetExportMode.IncludeMode;
                    exportCount = this.IncludeCount;
                }
                else if (this.ExcludeCount > 0)
                {
                    exportMode = BackupSetExportMode.ExcludeMode;
                    exportCount = this.ExcludeCount;
                }
                else
                {
                    exportMode = BackupSetExportMode.ExportAll;
                    exportCount = 0;
                }

                // Check to make sure we've got the go-ahead (if needed) to write out an export.
                if (isProcessingConfirmed == null)
                {
                    isProcessingConfirmed = ShouldProcess(workingCatalog.CatalogPath, res.SPCmdletExportSPBackupCatalog_ShouldProcess1);
                }

                // Carry out delete if we've gotten the thumbs-up
                if (isProcessingConfirmed.HasValue && isProcessingConfirmed == true)
                {
                    BackupCatalogUtilities.ExportBackupCatalog(workingCatalog, this.ExportPath, exportFileOrFolder, 
                        exportMode, exportCount, this.IgnoreBackupSetErrors, this.NoCompression, this.UpdateCatalogPath);
                    workingCatalog.Refresh();
                }
            }
            catch (Exception ex)
            {
                // Something went wrong with the backup catalog, the archiving process, or something in-between.
                // Regardless, we need to terminate the pipeline -- nothing more can be done.
                ThrowTerminatingError(ex, psError.Category, psError.Object);
            }

        }


        #endregion Overrides (SPSetCmdletBase)


        #region Properties


        /// <summary>
        /// This property is used to identify the number of most recent backup sets that should be
        /// excluded from the export operation. All other backup sets are included in the export
        /// that is performed. This property operates as a complement to the <see cref="IncludeCount"/>
        /// property.
        /// </summary>
        [Parameter(
            HelpMessage = "The total number of non error-containing backup sets (starting at the most recent backup set) that should be excluded from the export operation. All other backup sets are exported. For example, specifying an -ExcludeCount value of 3 will result in the export of all backup sets except for the three most recent non error-containing backup sets. If the -ExcludeCount parameter is specified, the -IncludeCount parameter cannot be specified; specifying both parameters will yield an error. Excluding both -ExcludeCount and -IncludeCount parameters will result in the export of the entire backup catalog. Note that the -ExcludeCount parameter is analogous to the -RetainCount parameter that is used with the Remove-SPBackupCatalog cmdlet.",
            Mandatory = false,
            Position = 3,
            ValueFromPipeline = false)]
        [ValidateRange(1, Int32.MaxValue)]
        public Int32 ExcludeCount { get; set; }


        /// <summary>
        /// Specifies the filename base (when compression is applied) or subdirectory name (when
        /// no compression is used) where exports will be created within the <see cref="ExportPath"/>
        /// </summary>
        [Parameter(
            HelpMessage = "This parameter is used to specify the base name of the export file (without extension) that is created in the location specified by the -ExportPath. If the -NoCompression switch is used, then this parameter specifies the name of the subdirectory that will house the exported backup sets in the location specified by the -ExportPath parameter. If the -ExportName parameter is not specified, a DateTime serial value is generated and used as either the archive file base name or subdirectory name as needed.",
            Mandatory = false,
            Position = 2,
            ValueFromPipeline = false)]
        public String ExportName { get; set; }
        

        /// <summary>
        /// This property is used to specify the location to which the export takes place. This is
        /// typically specified as a UNC path, and it's where the zip file gets created *or* a
        /// subdirectory is created (depending on whether or not compression is enabled).
        /// </summary>
        [Parameter(
            HelpMessage = "This path identifies the root file location where the export file(s) should be created. This parameter should normally be specified as a UNC path. In limited circumstances, such as with an all-in-one SharePoint server, a local path (beginning with a drive letter) may be acceptable.",
            Mandatory = true,
            Position = 1,
            ValueFromPipeline = false)]
        [ValidateNotNullOrEmpty]
        public String ExportPath { get; set; }

        
        /// <summary>
        /// This property represents the backup catalog that will be the target for export creation. Since the
        /// input is an <see cref="SPBackupCatalogPipeBind"/>, either a live <see cref="SPBackupCatalog"/>
        /// can be supplied - or some other object that can be coerced/interpreted by the pipe bind type.
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPBackupCatalogPipeBind"/>
        [Parameter(
            HelpMessage = "The backup catalog which will serve as the target for export selection. The backup catalog can be specified as an existing SPBackupCatalog object or as a path to the target backup catalog root (i.e., where the spbrtoc.xml resides).",
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true)]
        [ValidateNotNullOrEmpty]
        public SPBackupCatalogPipeBind Identity { get; set; }


        [Parameter(
            HelpMessage = "When -ExcludeCount and -IncludeCount selections are performed, they omit backup sets that have errors. This also means that if a full backup set contains errors and is omitted from export, then any differential backups that depend on the full backup set are also omitted. This is done to make operations equivalent to those of the Remove-SPBackupCatalog cmdlet. If the -IgnoreBackupSetErrors switch is specified, then -ExcludeCount and -IncludeCount selections are made without regard to whether or not backup sets contain errors.")]
        public SwitchParameter IgnoreBackupSetErrors { get; set; }


        /// <summary>
        /// This property is used to identify the number of oldest backup sets that should be
        /// explicitly included in the export operation. All other backup sets are excluded from the 
        /// export that is performed. This property operates as a complement to the <see cref="ExcludeCount"/>
        /// property.
        /// </summary>
        [Parameter(
            HelpMessage = "The total number of non error-containing backup sets (starting at the oldest backup set) that should be included in the export operation. All newer backup sets are excluded from the export. For example, specifying an -IncludeCount value of 2 will result in the export of the two oldest non error-containing backup sets. If the -IncludeCount parameter is specified, the -ExcludeCount parameter cannot be specified; specifying both parameters will yield an error. Excluding both -IncludeCount and -ExcludeCount parameters will result in the export of the entire backup catalog.",
            Mandatory = false,
            Position = 4,
            ValueFromPipeline = false)]
        [ValidateRange(1, Int32.MaxValue)]
        public Int32 IncludeCount { get; set; }


        [Parameter(
            HelpMessage = "When a backup export is created, one or more backup sets are normally compressed into a single zip archive. If this switch is used, then backup sets are \"loosely\" copied in uncompressed format to the destination folder specified by a combination of the -ExportPath parameter and the -ExportName parameter.")]
        public SwitchParameter NoCompression { get; set; }


        [Parameter(
            HelpMessage = "When this switch is set, the backup set pointers contained within the catalog will be updated to make them consistent with the -ExportPath (and the -ExportName when it is part of the destination path structure). This switch is commonly used in combination with the -NoCompression switch to make the exported backup sets usable in their destination location.")]
        public SwitchParameter UpdateCatalogPath { get; set; }


        #endregion Properties


    }
}
