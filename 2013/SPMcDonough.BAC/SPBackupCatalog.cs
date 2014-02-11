#region Namespace Imports


// Standard class imports
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// Needed for some folder and I/O related operations
using System.IO;

// Support for Farm-level types and extracting default backup location.
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Backup;

// For XML and LINQ processing
using System.Xml;
using System.Xml.Linq;

// For documentation generation using Gary Lapointe's MAML generator
using Lapointe.PowerShell.MamlGenerator.Attributes;

// Alias the resources
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{


    /// <summary>
    /// This class is used to represent one or more farm-level SharePoint 2010 backups. An instance of
    /// this object points to the root folder where backup sets are written to by <c>Backup-SPFarm</c> and
    /// Central Administration, and each location can contain a table of contents file and any number 
    /// of <c>spbrxxxx</c> subfolders that house discrete backup runs.
    /// </summary>
    public class SPBackupCatalog
    {


        #region Member Declarations


        // Backing property values
        private String _catalogPath;
        private Int64? _catalogSize;
        private Decimal? _sizePercentOfLastFull;

        // Internal tracking variables
        private XElement _lastBackupNode;
        private XElement _lastFullBackupNode;
        private XElement _tocFile;


        #endregion Member Declarations


        #region Constructors


        /// <summary>
        /// Default constructor; typically easier to simply use the overloaded constructor for <c>Path</c> assignment.
        /// </summary>
        public SPBackupCatalog()
        { }


        /// <summary>
        /// Overloaded constructor that allows <c>CatalogPath</c> assignment in one step.
        /// </summary>
        /// <param name="catalogPath">The file path to the root folder for the backup catalog.</param>
        public SPBackupCatalog(String catalogPath)
        {
            _catalogPath = catalogPath;
        }


        /// <summary>
        /// Overloaded constructor that infers <c>CatalogPath</c> from default backup location in farm.
        /// </summary>
        /// <param name="targetFarm">A reference to the <c>SPFarm</c> which will be used to extract a default
        /// backup location from.</param>
        public SPBackupCatalog(SPFarm targetFarm)
        {
            SPBackupRestoreConfigurationSettings defaultSettings = SPBackupRestoreConfigurationSettings.CreateSettings(targetFarm);
            _catalogPath = defaultSettings.PreviousBackupLocation;
        }


        #endregion Constructors


        #region Properties


        /// <summary>
        /// The file path to the root folder for the backup catalog. This is the folder that contains
        /// any number of <c>spbrxxxx</c> subfolders and an <c>spbrtoc.xml</c> table of contents file.
        /// </summary>
        /// <value>A <c>String</c> that represents the fully-qualified file path to the backup catalog root.</value>
        public String CatalogPath 
        {
            get { return _catalogPath; }
            set 
            {
                if (value != null && _catalogPath != value)
                { 
                    _catalogPath = value;
                    ResetProperties();
                }
            }
        }


        /// <summary>
        /// This read-only property returns the current size of the backup catalog, table of contents
        /// file, and all constituent backups in the <c>spbrxxxx</c> subfolders.
        /// </summary>
        /// <value>An <c>Int64</c> that represents the total size (in bytes) of the catalog and all
        /// backups contained within.</value>
        public Int64? CatalogSize
        {
            get
            {
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                } 
                if (!_catalogSize.HasValue && IsValidCatalog)
                {
                    _catalogSize = BackupCatalogUtilities.SumAllFileSizes(this.CatalogPath);
                }
                return _catalogSize;
            }
        }


        /// <summary>
        /// The total number of differential content + configuration backups housed in the catalog.
        /// This property does not report back configuration-only backups.
        /// </summary>
        /// <value>An <c>Int32</c> corresponding to the total number of differential content + configuration
        /// backups present. If an invalid <see cref="CatalogPath"/> is specified, a null value is returned.</value>
        public Int32? DifferentialBackupCount
        {
            get
            {
                Int32? diffCount = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (_tocFile != null)
                {
                    diffCount = _tocFile.Elements().Where(ho =>
                        (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "DIFFERENTIAL") &&
                         ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                         ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").Count();
                }

                return diffCount;
            }
        }


        /// <summary>
        /// The total number of full content + configuration backups housed in the catalog.
        /// This property does not report back configuration-only backups.
        /// </summary>
        /// <value>An <c>Int32</c> corresponding to the total number of full content + configuration
        /// backups present. If an invalid <see cref="CatalogPath"/> is specified, a null value is returned.</value>
        public Int32? FullBackupCount
        {
            get
            {
                Int32? fullCount = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (_tocFile != null)
                {
                    fullCount = _tocFile.Elements().Where(ho =>
                        (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                         ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                         ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").Count();
                }
                return fullCount;
            }
        }


        /// <summary>
        /// This read-only property is used to indicate whether or not an <c>spbrtoc.xml</c> table
        /// of contents file exists at the specified <see cref="CatalogPath"/>.
        /// </summary>
        /// <value>TRUE if a table of contents file exists at the specified <see cref="CatalogPath"/>,
        /// FALSE if not.</value>
        public Boolean IsValidCatalog
        {
            get
            {
                try
                {
                    return File.Exists(Path.Combine(this.CatalogPath, Globals.TOC_FILENAME));
                }
                catch
                {
                    return false;
                }
            }
        }



        /// <summary>
        /// The total number of errors that were encountered in the last full or differential backup.
        /// This property does not report back on configuration-only backups.
        /// </summary>
        /// <value>An <c>Int32</c> corresponding to the total number of errors identified in the last
        /// (non-config-only) full or differential backup. If an invalid <see cref="CatalogPath"/> is 
        /// specified, a null value is returned.</value>
        public Int32? LastBackupErrorCount
        {
            get
            {
                Int32? errorCount = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPErrorCount") != null)
                {
                    String errorNode = this.LastBackupNode.Element("SPErrorCount").Value;
                    if (errorNode != null)
                    {
                        errorCount = Convert.ToInt32(errorNode);
                    }
                }
                return errorCount;
            }
        }


        /// <summary>
        /// Provides the duration (length) of the last backup operation
        /// </summary>
        /// <value>A TimeSpan value that is simply a calculated value of the last backup
        /// finish time minus the last backup start time.</value>
        public TimeSpan? LastBackupDuration
        {
            get
            {
                TimeSpan? backupDuration = null;
                if (this.LastBackupFinish != null && this.LastBackupStart != null)
                {
                    backupDuration = this.LastBackupFinish.Value.Subtract(this.LastBackupStart.Value);
                }
                return backupDuration;
            }
        }


        /// <summary>
        /// Provides the date and time at which the last backup completed.
        /// </summary>
        /// <value>A DateTime value indicating finishing date/time for last backup.</value>
        public DateTime? LastBackupFinish
        {
            get
            {
                DateTime? backupFinish = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPFinishTime") != null)
                {
                    String finishTimeNode = this.LastBackupNode.Element("SPFinishTime").Value;
                    if (finishTimeNode != null)
                    {
                        backupFinish = Convert.ToDateTime(finishTimeNode);
                    }
                }
                return backupFinish;
            }
        }


        /// <summary>
        /// Indicates which backup method, "Full" or "Differential", was used to create the last
        /// non-configuration-only backup in the backup catalog.
        /// </summary>
        /// <value>A String that contains the backup method used: either "Full" or "Differential"</value>
        public String LastBackupMethod
        {
            get
            {
                String backupMethod  = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPBackupMethod") != null)
                {
                    backupMethod = this.LastBackupNode.Element("SPBackupMethod").Value;
                }
                return backupMethod;
            }
        }


        /// <summary>
        /// This internal property is used to expose the last backup node for additional actions.
        /// The last node is defined as the most recent SPHistoryObject that is not a configuration-only
        /// backup
        /// </summary>
        internal XElement LastBackupNode
        {
            get
            {
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (_lastBackupNode == null && IsValidCatalog)
                {
                    var revSort = _tocFile.Elements().Where(ho =>
                        ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                        ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").OrderByDescending(fho =>
                            Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
                    _lastBackupNode = revSort.FirstOrDefault();
                }
                return _lastBackupNode;
            }
        }


        /// <summary>
        /// Indicates the full path that was used for the last backup that was created.
        /// </summary>
        /// <value>A String that contains the user account (e.g., SPDC\Administrator)</value>
        /// <remarks>This path may not directly match a subdirectory under the current backup
        /// catalog; e.g., if a local path spec is used for the current backup catalog but the
        /// last backup operation was conducted against the same location through a UNC path.</remarks>
        public String LastBackupPath
        {
            get
            {
                String backupPath = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPBackupDirectory") != null)
                {
                    backupPath = this.LastBackupNode.Element("SPBackupDirectory").Value;
                }
                return backupPath;
            }
        }


        /// <summary>
        /// Identifies the account of the user who requested the last backup.
        /// </summary>
        /// <value>A String that contains the user account (e.g., SPDC\Administrator)</value>
        public String LastBackupRequestor
        {
            get
            {
                String backupRequestor = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPRequestedBy") != null)
                {
                    backupRequestor = this.LastBackupNode.Element("SPRequestedBy").Value;
                }
                return backupRequestor;
            }
        }


        /// <summary>
        /// This property returns the size (in bytes) of the last backup set.
        /// </summary>
        public Int64? LastBackupSize
        {
            get
            {
                Int64? backupSize = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPDirectoryName") != null)
                {
                    try
                    {
                        // Combine the relative path with our base catalog path to get the full path that
                        // we'll evaluate.
                        String relativePath = this.LastBackupNode.Element("SPDirectoryName").Value;
                        String backupSetPath = Path.Combine(this.CatalogPath, relativePath);
                        backupSize = BackupCatalogUtilities.SumAllFileSizes(backupSetPath);
                    }
                    catch
                    {
                        backupSize = null;
                    }
                }
                return backupSize;
            }
        }


        /// <summary>
        /// Provides the date and time at which the last backup kicked-off.
        /// </summary>
        /// <value>A DateTime value indicating start date/time for last backup.</value>
        public DateTime? LastBackupStart
        {
            get
            {
                DateTime? backupStart = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPStartTime") != null)
                {
                    String startTimeNode = this.LastBackupNode.Element("SPStartTime").Value;
                    if (startTimeNode != null)
                    {
                        backupStart = Convert.ToDateTime(startTimeNode);
                    }
                }
                return backupStart;
            }
        }


        /// <summary>
        /// Identifies the top-most component (in the farm hierarchy) in the last backup run.
        /// </summary>
        /// <value>A String that contains the top-most component (e.g., "Farm")</value>
        public String LastBackupTopComponent
        {
            get
            {
                String topComponent = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPTopComponent") != null)
                {
                    topComponent = this.LastBackupNode.Element("SPTopComponent").Value;
                }
                return topComponent;
            }
        }


        /// <summary>
        /// The total number of warnings that were encountered in the last full or differential backup.
        /// This property does not report back on configuration-only backups.
        /// </summary>
        /// <value>An <c>Int32</c> corresponding to the total number of warnings identified in the last
        /// (non-config-only) full or differential backup. If an invalid <see cref="CatalogPath"/> is 
        /// specified, a null value is returned.</value>
        public Int32? LastBackupWarningCount
        {
            get
            {
                Int32? warningCount = null;
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (this.LastBackupNode != null && this.LastBackupNode.Element("SPWarningCount") != null)
                {
                    String warningNode = this.LastBackupNode.Element("SPWarningCount").Value;
                    if (warningNode != null)
                    {
                        warningCount = Convert.ToInt32(warningNode);
                    }
                }
                return warningCount;
            }
        }


        /// <summary>
        /// This internal property is used to expose the last full backup node for additional actions.
        /// The last full node is defined as the most recent SPHistoryObject that is not a configuration-only
        /// backup and has a backup method of "Full"
        /// </summary>
        internal XElement LastFullBackupNode
        {
            get
            {
                if (_tocFile == null && IsValidCatalog)
                {
                    Refresh();
                }
                if (_lastFullBackupNode == null && IsValidCatalog)
                {
                    var revSort = _tocFile.Elements().Where(ho =>
                        ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                        ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL" &&
                        ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").OrderByDescending(fho =>
                            Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
                    _lastFullBackupNode = revSort.FirstOrDefault();
                }
                return _lastFullBackupNode;
            }
        }


        /// <summary>
        /// This property will return the last backup set's size as a percentage of the previous full backup that was taken 
        /// This value can be useful if backup scripts are trying to determine if differential backups are getting too big
        /// and a new baseline (full backup) should be taken instead.
        /// </summary>
        /// <value>An <c>Decimal</c> corresponding to the size of the last backup relative to the previous full backup. For
        /// differential backups, this value will usually be less than 100. For full backups, values will range below and
        /// above 100.</value>
        public Decimal? SizePercentageOfLastFull
        {
            get
            {
                if (!_sizePercentOfLastFull.HasValue && IsValidCatalog && this.LastBackupNode != null && this.LastFullBackupNode != null &&
                    this.LastBackupNode.Element("SPDirectoryName") != null &&
                    this.LastFullBackupNode.Element("SPDirectoryName") != null)
                {
                    // Build the paths to the two folders for comparison
                    String lastFullBackupPath = Path.Combine(_catalogPath, this.LastFullBackupNode.Element("SPDirectoryName").Value);
                    String lastBackupPath = Path.Combine(_catalogPath, this.LastBackupNode.Element("SPDirectoryName").Value);

                    // Get sizes for each of the folders.
                    Int64? bytesInLastFullBackup = BackupCatalogUtilities.SumAllFileSizes(lastFullBackupPath);
                    Int64? bytesInLastBackup = BackupCatalogUtilities.SumAllFileSizes(lastBackupPath);

                    // Divide and multiply by 100 to get the percentage.
                    if (bytesInLastBackup.HasValue && bytesInLastFullBackup.HasValue)
                    {
                        Double divisionResult = (Convert.ToDouble(bytesInLastBackup.Value) / Convert.ToDouble(bytesInLastFullBackup.Value)) * 100;
                        _sizePercentOfLastFull = Decimal.Round(Convert.ToDecimal(divisionResult), 2);
                    }
                }
                return _sizePercentOfLastFull;
            }
        }


        /// <summary>
        /// This property provides an easy way for internal types to get at the table
        /// of contents data for the backup catalog and make changes to it. If changes are made,
        /// they should be persisted to disk with a call to <see cref="SaveTableOfContents"/>
        /// </summary>
        /// <value>An <x>XElement</x> that contains the table of contents XML data for the
        /// backup catalog.</value>
        /// <seealso cref="SaveTableOfContents"/>
        internal XElement TableOfContents
        {
            get { return _tocFile; }
            set
            {
                _tocFile = value;
            }
        }


        #endregion Properties


        #region Methods (Public)


        /// <summary>
        /// This method is responsible for performing a property clear and re-fetch of the TOC
        /// file contents to ensure that properties are up-to-date.
        /// </summary>
        public void Refresh()
        {
            // Reset property values.
            ResetProperties();

            // Load the table of contents file if possible. If the file isn't present or doesn't load,
            // simply use a null to indicate problems.
            try
            {
                _tocFile = XElement.Load(Path.Combine(this.CatalogPath, Globals.TOC_FILENAME));
            }
            catch
            {
                _tocFile = null;
            }
        }


        /// <summary>
        /// This method is used to persist the backup catalog's table of contents file back to disk.
        /// This is typically done after assignment has been performed (through the <see cref="TableOfContents"/>
        /// property) of new TOC XML.
        /// </summary>
        /// <seealso cref="TableOfContents"/>
        public void SaveTableOfContents()
        {
            String fullPath = Path.Combine(this.CatalogPath, Globals.TOC_FILENAME);
            _tocFile.Save(fullPath);
        }


        #endregion Methods (Public)


        #region Methods (Private)


        /// <summary>
        /// This method is responsible for "zeroing out" a catalog instance. As new properties and internal
        /// tracking members are added, this method needs to be modified to consider them.
        /// </summary>
        private void ResetProperties()
        {
            _catalogSize = null;
            _sizePercentOfLastFull = null;
            _lastBackupNode = null;
            _lastFullBackupNode = null;
            _tocFile = null;
        }
        

        #endregion Methods (Private)


    }
}
