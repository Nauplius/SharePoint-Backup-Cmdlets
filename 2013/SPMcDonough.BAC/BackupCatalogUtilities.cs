#region Namespace Imports


// Standard class imports
using System;
using System.Collections.Generic;
using System.Linq;

// Needed for some folder and I/O related operations
using System.IO;

// For XML and LINQ processing
using System.Xml.Linq;

// For working with Zip files using the DotNetZip Library (which is available
// from CodePlex at http://dotnetzip.codeplex.com).
using Ionic.Zip;

// Set resources alias
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{

    /// <summary>
    /// This class contains the methods that are actually used to interact with the backup
    /// catalogs (as <see cref="SPBackupCatalog"/> objects) and files that they contain.
    /// </summary>
    /// <seealso cref="SPBackupCatalog"/>
    internal static class BackupCatalogUtilities
    {

        #region Methods (Public, Static)

        /// <summary>
        /// This method is used to examine an <see cref="SPBackupCatalog"/> and build a String that can
        /// be used as the body of an e-mail to communicate the status of the last backup run.
        /// </summary>
        /// <param name="workingCatalog">An <see cref="SPBackupCatalog"/> that will be analyzed for
        /// purposes of building the status message.</param>
        /// <returns>A string containing HTML that can be used as the body of an e-mail message.</returns>
        public static String BuildEmailStatusBody(SPBackupCatalog workingCatalog)
        {
            // Build a list for argument insertion into the template we'll eventually pull from the
            // resources area.
            var mailArgs = new List<Object>
            {
                DateTime.Now.ToShortTimeString(),
                DateTime.Now.ToLongDateString(),
                workingCatalog.LastBackupPath,
                workingCatalog.LastBackupTopComponent,
                workingCatalog.LastBackupMethod,
                workingCatalog.LastBackupRequestor
            };

            try
            {
                mailArgs.Add(Globals.IntelligentDateTimeFormat(workingCatalog.LastBackupStart));
            }
            catch (FormatException)
            {
                mailArgs.Add(null);
            }

            // The first two arguments are the time {0} and date {1}

            // The next block is for general backup information: location {2}, top component {3},
            // backup method {4}, and requestor {5}

            // The final block is for start date/time {6}, finish date/time {7}, duration {8},
            // backup set size {9}, error count {10}, and warning count {11}

            try
            {
                mailArgs.Add(Globals.IntelligentDateTimeFormat(workingCatalog.LastBackupFinish));
            }
            catch (FormatException)
            {
                mailArgs.Add(null);
            }

            try
            {
                mailArgs.Add(Globals.IntelligentTimeFormat(workingCatalog.LastBackupDuration));
            }
            catch (FormatException)
            {
                mailArgs.Add(null);
            }
            
            mailArgs.Add(Globals.IntelligentBytesFormat(workingCatalog.LastBackupSize));
            mailArgs.Add(Globals.IntelligentIntegerFormat(workingCatalog.LastBackupErrorCount));
            mailArgs.Add(Globals.IntelligentIntegerFormat(workingCatalog.LastBackupWarningCount));

            // Execute a replacement against the template using the argument list we built up and return
            // it for e-mail purposes.
            var bodyTemplate = res.Send_SPBackupStatus_Status_Template;
            var paramVals = mailArgs.ToArray();
            var mailBody = String.Format(bodyTemplate, paramVals);
            return mailBody;
        }


        /// <summary>
        /// This method is called to clear the contents of an <see cref="SPBackupCatalog"/> to remove
        /// all backup sets it contains. This operation ensures that an spbrtoc.xml file remains with
        /// any remnants of restore and non-backup operations, but backup set references (and folders
        /// containing the sets) are deleted.
        /// <para>"Removal" of a backup set involves two different operations. The first operation is
        /// the actual updating of the table of contents file to reflect that the backup set is no
        /// longer present. The second part of the removal is the actual deletion of the backup
        /// subfolders housing the backup data being trimmed out.</para>
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPCmdletRemoveSPBackupCatalog"/>
        /// <param name="workingCatalog">An <see cref="SPBackupCatalog"/> object that represents the
        /// backup set to be cleared.</param>
        /// <remarks>If an exception occurs, it is assumed that the calling method will trap it and
        /// handle it accordingly (including a possible re-throw).</remarks>
        public static void EmptyBackupCatalog(SPBackupCatalog workingCatalog)
        {
            // Use the XML for the backup catalog's TOC to grab all of the full backup SPHistoryObject
            // nodes and reverse sort them by directory number.
            XElement tocFile = workingCatalog.TableOfContents;
            String backupRoot = workingCatalog.CatalogPath;
            var fullHistoryObjects = tocFile.Elements().Where(ho =>
                (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                 ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                 ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").OrderByDescending(fho =>
                    Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));

            // We need each of the folders we'll be deleting as part of the operation. This has to be
            // dumped to a list because subsequent deferred evaluation (when deleting folders) wouldn't
            // work properly following PHASE 1 (below).
            var backupSetsToDelete = tocFile.Elements().Where(ho =>
                (ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE")).Select(dho =>
                new { ID = dho.Element("SPId").Value, Folder = Path.Combine(backupRoot, dho.Element("SPDirectoryName").Value) }).ToList();

            // CLEANUP PHASE 1: clean-up the SPHistoryObject elements in the backup TOC file and save off the file
            tocFile.Elements().Where(ho =>
                (ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE")).Remove();
            workingCatalog.TableOfContents = tocFile;
            workingCatalog.SaveTableOfContents();

            // CLEANUP PHASE 2: Iterate through the backup folders to delete and clear them out.
            foreach (var backupSet in backupSetsToDelete)
            {
                Directory.Delete(backupSet.Folder, true);
            }
        }


        /// <summary>
        /// This method is called to carry out an export of backup sets contained in a backup catalog.
        /// In essence, this method does all of the heavy lifting for the Export-SPBackupCatalog cmdlet
        /// housed in the <see cref="SPCmdletExportSPBackupCatalog"/> class. All parameters for this 
        /// method map in some way to acceptable parameters for the cmdlet.
        /// </summary>
        /// <seealso cref="SPCmdletExportSPBackupCatalog"/>
        public static void ExportBackupCatalog(SPBackupCatalog workingCatalog, String exportPath, String exportName, 
            BackupSetExportMode exportMode, Int32 exportCount, Boolean ignoreErrors, Boolean useNoCompression,
            Boolean updateCatalogPath)
        {
            // Use the XML for the backup catalog's TOC to grab all of the full backup SPHistoryObject
            // nodes and reverse sort them by directory number.
            XElement tocFile = workingCatalog.TableOfContents;
            IOrderedEnumerable<XElement> fullHistoryObjects;

            // If we aren't ignoring errors, then we'll need to limit (potentially) the fullHistoryObjects
            // needed to just those backup sets that don't contain errors.
            if (ignoreErrors)
            {
                // All SPHistoryObjects are valid
                fullHistoryObjects = tocFile.Elements().Where(ho =>
                (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                    ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                    ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").OrderByDescending(fho =>
                    Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
            }
            else
            {
                // Only SPHistoryObjects with a zero error count are valid
                fullHistoryObjects = tocFile.Elements().Where(ho =>
                (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                    ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                    ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                    Convert.ToInt32(ho.Element("SPErrorCount").Value) == 0).OrderByDescending(fho =>
                    Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
            }

            // We've got the appropriate backup sets selected; now we may need to narrow them down based
            // on the type export action/mode specified.
            var lastKeepIndex = 0;
            var lastDirNumber = -1;

            switch (exportMode)
            {
                case BackupSetExportMode.ExportAll:
                    // Everything is being exported, so our keep index is out of bounds
                    lastKeepIndex = -1;
                    break;
                case BackupSetExportMode.IncludeMode:
                    // We're explicitly including backup sets, so we need to determine the last directory
                    // for export by going to the end and count up the list to find the last one to keep.
                    lastKeepIndex = fullHistoryObjects.Count() - exportCount - 1;
                    break;
                case BackupSetExportMode.ExcludeMode:
                    // Get the directory number of the last full backup we want to export; anything below that
                    // will be ignored during export.
                    lastKeepIndex = exportCount - 1;
                    break;
            }

            // Determine what we're going to actually export based on keep indexes and directory numbers.
            // If the index is less than zero
            if (lastKeepIndex >= (fullHistoryObjects.Count() - 1))
            {
                // We aren't actually exporting any objects, so select a directory number that we know
                // none of the backup sets will be below.
                lastDirNumber = Int32.MinValue;
            }
            else if (lastKeepIndex >= 0)
            {
                // We're exporting some and not exporting others. Locate the element at the appropriate
                // index and figure out what its directory number is. Everything under that number will
                // be what we export.
                lastDirNumber = Convert.ToInt32(fullHistoryObjects.ElementAt(lastKeepIndex).Element("SPDirectoryNumber").Value);
            }
            else
            {
                // We have an index below zero. This means that *everything* is going to be exported.
                lastDirNumber = Int32.MaxValue;
            }

            // We've got the directory number of the backup set which marks the start of everything we
            // *don't* want to export. We now need to select all backup sets have a directory number that is
            // *less* than the last directory number and build folder paths.
            //
            // NOTE: Only full and differential content backups are included for export. Config-only backups
            // are excluded. The ignoreErrors switch is also obeyed to determine if backups with errors are
            // included in the export.
            IEnumerable<XElement> exportSelections = tocFile.Elements().Where(ho =>
                (ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                 ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                 Convert.ToInt32(ho.Element("SPDirectoryNumber").Value) < lastDirNumber));

            // If we aren't ignoring errors, then we have some clean-up to do. Differential backups themselves 
            // are easy to drop (no dependencies), but we'll need some special logic to clean up full backups
            // because of the fact that differentials may be tied to them.
            var additionalDirectoryExclusions = new List<Int32>();
            if (!ignoreErrors)
            {
                // Work our way "backward" through the collection. Oldest sets are at the end of the collection,
                // so if we encounter a full backup with errors we'll need to delete the differentials that are
                // tied to it.
                Boolean isBadFullBackup = false;
                for (Int32 backupIndex = exportSelections.Count() - 1; backupIndex >= 0; backupIndex--)
                {
                    // Grab a reference to the current backup and determine if it is full or differential.
                    XElement currentBackup = exportSelections.ElementAt(backupIndex);
                    if (currentBackup.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL")
                    {
                        // This could be a lone full backup, or we may be at the start of a sequence that includes
                        // any number of differential backups. Determine if the current backup has errors.
                        if (Convert.ToInt32(currentBackup.Element("SPErrorCount").Value) > 0)
                        {
                            // There are one or more errors in the full backup. It will not only need to be excluded,
                            // but we'll also have to flip the "bad backup" flag so subsequent differentials are 
                            // added to the exclusions list.
                            additionalDirectoryExclusions.Add(Convert.ToInt32(currentBackup.Element("SPDirectoryNumber").Value));
                            isBadFullBackup = true;
                        }
                        else
                        {
                            // No errors in the full backup set; we're good to go for this backup and any
                            // differentials that are tied to it.
                            isBadFullBackup = false;
                        }
                    }
                    else
                    {
                        // We're working with a differential backup. If it has an error, or if the full backup that
                        // it is tied to have an issue, we need to flag it for deletion.
                        if (isBadFullBackup || Convert.ToInt32(currentBackup.Element("SPErrorCount").Value) > 0)
                        {
                            // Add the current differential to the exclusion list.
                            additionalDirectoryExclusions.Add(Convert.ToInt32(currentBackup.Element("SPDirectoryNumber").Value));
                        }
                    }
                }
            }

            // If we've actually got one or more backup sets to export, we can continue with file operations.
            if (exportSelections.Count() > 0)
            {
                // We've got a selection of backup sets for export. We need to transform them into a list
                // that we can use for the actual export operations.
                //
                // We'll be cross-checking to make sure that the directory number of each backup isn't present in
                // the exclude list before selecting it for backup.
                String backupRoot = workingCatalog.CatalogPath;
                var backupSetsToExport = exportSelections.Where(bse => 
                    additionalDirectoryExclusions.IndexOf(Convert.ToInt32(bse.Element("SPDirectoryNumber").Value)) == -1).Select(se =>
                    new { RelativeFolder = se.Element("SPDirectoryName").Value, SourceFolder = Path.Combine(backupRoot, se.Element("SPDirectoryName").Value) }).ToList();

                // Build-out the new table-of-contents file for the export we're about to perform (and again,
                // observe exclusions that have been added to the appropriate list).
                XElement newTocFile = XElement.Parse(Globals.BACKUP_CATALOG_BASE_XML);
                foreach (XElement currentHistoryObject in exportSelections.Where(bse =>
                    additionalDirectoryExclusions.IndexOf(Convert.ToInt32(bse.Element("SPDirectoryNumber").Value)) == -1))
                {
                    newTocFile.Add(currentHistoryObject);
                }

                // If we need to do path corrections, now is the time.
                if (updateCatalogPath)
                {
                    // We're going to update the catalog path, so we need to determine what the new path will be.
                    // The actual path depends on whether or not the ExportName is part of the path.
                    String newCatalogPath;
                    if (useNoCompression)
                    {
                        newCatalogPath = Path.Combine(exportPath, exportName);
                    }
                    else
                    {
                        newCatalogPath = exportPath;
                    }

                    // Execute the update.
                    newTocFile = UpdateBackupCatalogPaths(newTocFile, newCatalogPath);
                }

                // How we proceed at this point depends on whether or not compression is enabled.
                if (useNoCompression)
                {
                    // We're not using compression, so we need to create a folder structure at the export path.
                    // Build the full path and use it for the creation.
                    var exportFullPath = Path.Combine(exportPath, exportName);
                    Directory.CreateDirectory(exportFullPath);

                    // Write out the table of contents to export area.
                    var tocFilePath = Path.Combine(exportFullPath, Globals.TOC_FILENAME);
                    newTocFile.Save(tocFilePath);

                    // Iterate through each of the backup sets to export and copy them to the export area.
                    foreach (var exportBackupSet in backupSetsToExport)
                    {
                        // Create the destination folder.
                        var destinationFolder = Path.Combine(exportFullPath, exportBackupSet.RelativeFolder);
                        Directory.CreateDirectory(destinationFolder);

                        // We now need to recursively copy the contents of the source folder (and its subdirectories
                        // if any exist) into the destination folder.
                        RecursiveCopy(exportBackupSet.SourceFolder, destinationFolder);
                    }
                }
                else
                {
                    // We're using compression, so the export name supplied will be used as the name of the
                    // zip file we create. Make sure the export name actually ends in .zip
                    if (!exportName.EndsWith(Globals.EXPORT_ARCHIVE_EXTENSION, StringComparison.InvariantCultureIgnoreCase))
                    { exportName += Globals.EXPORT_ARCHIVE_EXTENSION; }
                    
                    // Build the full path to the zip file we want to create; if it already exists, delete it so
                    // we can write out the new one.
                    var archiveFullPath = Path.Combine(exportPath, exportName);
                    if (File.Exists(archiveFullPath))
                    { File.Delete(archiveFullPath); }

                    // Prep the zip archive that we'll be writing to
                    using (var catalogArchive = new ZipFile(archiveFullPath))
                    {
                        // Begin by writing out the table of contents to the root of the archive.
                        catalogArchive.AddEntry(Globals.TOC_FILENAME, newTocFile.ToString());

                        // Iterate through each of the backup sets to process them into the archive.
                        foreach (var exportBackupSet in backupSetsToExport)
                        {
                            catalogArchive.AddDirectory(exportBackupSet.SourceFolder, exportBackupSet.RelativeFolder);
                        }

                        // We need to save the archive, but we should make a couple of changes to ensure proper
                        // .zip processing. The archive that we're creating is likely very large relative to the
                        // average .zip file. For this reason, we'll adjust the save options and buffers based
                        // on a sample from the Ionic Zip Library "BufferSize" property description.
                        catalogArchive.UseZip64WhenSaving = Zip64Option.Always;
                        catalogArchive.BufferSize = 65536 * 8;      // 512k buffer (8k buffer is default)
                        catalogArchive.Save();
                    }
                }
            }
        }


        /// <summary>
        /// This method starts at a file path specified by <paramref name="rootFolder"/> and retrieves
        /// all of the files it contains. The sizes of the files are then summed-up and returned.
        /// </summary>
        /// <param name="rootFolder">A fully-qualified file path to somewhere in the local file system
        /// or a network file path</param>
        /// <returns>The total size (in bytes) of all of the files contained in the <paramref name="rootFolder"/>
        /// and its child folders. If the file path isn't valid or a problem occurs while iterating through
        /// the files, a null is returned.</returns>
        public static Int64? SumAllFileSizes(String rootFolder)
        {
            Int64? operationResult = 0;
            try
            {
                var allFileNames = Directory.GetFiles(rootFolder, "*", SearchOption.AllDirectories);
                var allFileInfos = allFileNames.Select(cf => new FileInfo(cf));
                operationResult = (Int64?)allFileInfos.Sum(fi => fi.Length);
            }
            catch
            {
                operationResult = null;
            }
            return operationResult;
        }
        
        
        /// <summary>
        /// This method is called to trim the contents of an <see cref="SPBackupCatalog"/> to retain
        /// only a certain number of full backups. The method is responsible for ensuring that only
        /// the most recent number of full backups specified are retained. Older full backups (and any
        /// differential backups that depend on those full backups) are removed from the backup set.
        /// <para>"Removal" of a backup set involves two different operations. The first operation is
        /// the actual updating of the table of contents file to reflect that the backup set is no
        /// longer present. The second part of the removal is the actual deletion of the backup
        /// subfolders housing the backup data being trimmed out.</para>
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPCmdletRemoveSPBackupCatalog"/>
        /// <param name="workingCatalog">An <see cref="SPBackupCatalog"/> object that represents the
        /// backup set to be trimmed.</param>
        /// <param name="retainCount">An <c>Int32</c> that contains the total number of full backups
        /// to retain in the <paramref name="workingCatalog"/>.</param>
        /// <param name="ignoreErrors">TRUE if backup set errors should be ignored for purposes of
        /// determining what to delete, FALSE if backup sets with errors should not count against the
        /// number of backups retained.</param>
        /// <remarks>If an exception occurs, it is assumed that the calling method will trap it and
        /// handle it accordingly (including a possible re-throw).</remarks>
        public static void TrimBackupCatalogCount(SPBackupCatalog workingCatalog, Int32 retainCount, Boolean ignoreErrors)
        {
            // Use the XML for the backup catalog's TOC to grab all of the full backup SPHistoryObject
            // nodes and reverse sort them by directory number.
            var tocFile = workingCatalog.TableOfContents;
            IOrderedEnumerable<XElement> fullHistoryObjects;

            // If we aren't ignoring errors, then we'll need to limit (potentially) the fullHistoryObjects
            // needed to just those backup sets that don't contain errors.
            if (ignoreErrors)
            {
                // All SPHistoryObjects are valid
                fullHistoryObjects = tocFile.Elements().Where(ho =>
                (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                    ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                    ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").OrderByDescending(fho =>
                    Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
            }
            else 
            {
                // Only SPHistoryObjects with a zero error count are valid
                fullHistoryObjects = tocFile.Elements().Where(ho =>
                (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                    ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                    ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                    Convert.ToInt32(ho.Element("SPErrorCount").Value) == 0).OrderByDescending(fho =>
                    Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
            }

            // If we've got more full history objects than our count, then it's time to do some trimming.
            if (fullHistoryObjects.Count() > retainCount)
            {
                // We now have all of the SPHistoryObject nodes that correspond to full backups, and they
                // are sorted from most recent to oldest. We want to keep the most recent ones (number
                // specified by retainCount) and dump everything else.
                //
                // Get the directory number of the last full backup we want to keep; anything below that
                // will be deleted. We'll also get a quick reference to the backup root folder.
                Int32 lastDirNumber = Convert.ToInt32(fullHistoryObjects.ElementAt(retainCount - 1).Element("SPDirectoryNumber").Value);
                String backupRoot = workingCatalog.CatalogPath;

                // We need each of the folders we'll be deleting as part of the operation. This has to be
                // dumped to a list because subsequent deferred evaluation (when deleting folders) wouldn't
                // work properly following PHASE 1 (below).
                var backupSetsToDelete = tocFile.Elements().Where(ho =>
                    (ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                    Convert.ToInt32(ho.Element("SPDirectoryNumber").Value) < lastDirNumber)).Select(dho =>
                    new { ID = dho.Element("SPId").Value, Folder = Path.Combine(backupRoot, dho.Element("SPDirectoryName").Value) }).ToList();

                // CLEANUP PHASE 1: clean-up the SPHistoryObject elements in the backup TOC file and save off the file
                tocFile.Elements().Where(ho =>
                    (ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                    Convert.ToInt32(ho.Element("SPDirectoryNumber").Value) < lastDirNumber)).Remove();
                workingCatalog.TableOfContents = tocFile;
                workingCatalog.SaveTableOfContents();

                // CLEANUP PHASE 2: Iterate through the backup folders to delete and clear them out.
                foreach (var backupSet in backupSetsToDelete)
                {
                    Directory.Delete(backupSet.Folder, true);
                }
            }
        }


        /// <summary>
        /// This method is called to trim the contents of an <see cref="SPBackupCatalog"/> to retain
        /// as many backups as possible that allow the backup catalog to remain below the retention
        /// size specified by the <paramref name="retainSize"/> parameter. Older full backups (and any
        /// differential backups that depend on those full backups) are removed from the backup set
        /// from oldest to newest, and the overall size of the backup set is checked with each removal.
        /// If the backup catalog is too big after a removal or set of removals, the process is repeated.
        /// <para>"Removal" of a backup set involves two different operations. The first operation is
        /// the actual updating of the table of contents file to reflect that the backup set is no
        /// longer present. The second part of the removal is the actual deletion of the backup
        /// subfolders housing the backup data being trimmed out. As indicated, this process is
        /// carried out recursively until the size constraints are met.</para>
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPCmdletRemoveSPBackupCatalog"/>
        /// <param name="workingCatalog">An <see cref="SPBackupCatalog"/> object that represents the
        /// backup set to be trimmed.</param>
        /// <param name="retainSize">An <c>Int64</c> that contains the total maximum size of all backups
        /// to retain in the <paramref name="workingCatalog"/>.</param>
        /// <param name="ignoreErrors">TRUE if backup set errors should be ignored for purposes of
        /// determining what to delete, FALSE if backup sets with errors should not count against the
        /// number of backups retained.</param>
        /// <remarks>If an exception occurs, it is assumed that the calling method will trap it and
        /// handle it accordingly (including a possible re-throw).</remarks>
        public static void TrimBackupCatalogSize(SPBackupCatalog workingCatalog, Int64 retainSize, Boolean ignoreErrors)
        {
            // Use the XML for the backup catalog's TOC to grab all of the full backup SPHistoryObject
            // nodes and reverse sort them by directory number.
            var backupRoot = workingCatalog.CatalogPath;
            var tocFile = workingCatalog.TableOfContents;
            IOrderedEnumerable<XElement> fullHistoryObjects;

            // If we aren't ignoring errors, then we'll need to limit (potentially) the fullHistoryObjects
            // needed to just those backup sets that don't contain errors.
            if (ignoreErrors)
            {
                // All SPHistoryObjects are valid
                fullHistoryObjects = tocFile.Elements().Where(ho =>
                (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                    ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                    ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").OrderByDescending(fho =>
                    Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
            }
            else
            {
                // Only SPHistoryObjects with a zero error count are valid
                fullHistoryObjects = tocFile.Elements().Where(ho =>
                (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                    ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                    ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                    Convert.ToInt32(ho.Element("SPErrorCount").Value) == 0).OrderByDescending(fho =>
                    Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
            }

            // We need to enter into a repeating delete cycle until one of two conditions are met:
            // 1. We have only one full backup set (plus any number of differentials) left
            // 2. The backup catalog is equal to our below the retention size supplied
            while (fullHistoryObjects.Count() > 1 && workingCatalog.CatalogSize > retainSize)
            {
                // Get the ID of the next-to-the-last full backup set; i.e., the last full backup
                // set that will remain in the catalog after the delete. Everything below it in the
                // directory number sequence will be wiped.
                var lastDirNumber = Convert.ToInt32(fullHistoryObjects.ElementAt(fullHistoryObjects.Count() - 2).Element("SPDirectoryNumber").Value);

                // We need each of the folders we'll be deleting as part of the operation. This has to be
                // dumped to a list because subsequent deferred evaluation (when deleting folders) wouldn't
                // work properly following PHASE 1 (below).
                var backupSetsToDelete = tocFile.Elements().Where(ho =>
                    (ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                    Convert.ToInt32(ho.Element("SPDirectoryNumber").Value) < lastDirNumber)).Select(dho =>
                    new { ID = dho.Element("SPId").Value, Folder = Path.Combine(backupRoot, dho.Element("SPDirectoryName").Value) }).ToList();

                // CLEANUP PHASE 1: clean-up the SPHistoryObject elements in the backup TOC file and save off the file
                tocFile.Elements().Where(ho =>
                    (ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                    Convert.ToInt32(ho.Element("SPDirectoryNumber").Value) < lastDirNumber)).Remove();
                workingCatalog.TableOfContents = tocFile;
                workingCatalog.SaveTableOfContents();

                // CLEANUP PHASE 2: Iterate through the backup folders to delete and clear them out.
                foreach (var backupSet in backupSetsToDelete)
                {
                    Directory.Delete(backupSet.Folder, true);
                }

                // Refresh the working catalog and re-query for the needed objects.
                workingCatalog.Refresh();
                if (ignoreErrors)
                {
                    // All SPHistoryObjects are valid
                    fullHistoryObjects = tocFile.Elements().Where(ho =>
                    (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                        ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                        ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE").OrderByDescending(fho =>
                        Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
                }
                else
                {
                    // Only SPHistoryObjects with a zero error count are valid
                    fullHistoryObjects = tocFile.Elements().Where(ho =>
                    (ho.Element("SPBackupMethod").Value.Trim().ToUpper() == "FULL") &&
                        ho.Element("SPConfigurationOnly").Value.Trim().ToUpper() == "FALSE" &&
                        ho.Element("SPIsBackup").Value.Trim().ToUpper() == "TRUE" &&
                        Convert.ToInt32(ho.Element("SPErrorCount").Value) == 0).OrderByDescending(fho =>
                        Convert.ToInt32(fho.Element("SPDirectoryNumber").Value));
                }
            }
        }


        #region UpdateBackupCatalogPaths


        /// <summary>
        /// This method is called when the backup set pointers contained within <paramref name="workingCatalog"/>
        /// need to be validated and corrected to make them consistent with the current path held by 
        /// the SPBackupCatalog. This type of correction becomes necessary when a backup catalog that 
        /// was created at one path (for example, "\\BackupServer\Backups") is moved to another path 
        /// (for example, "\\BackupServer\Restores") and referenced. 
        /// <para>In this scenario, folder pointers that are internal to the backup catalog will still 
        /// point to "\\BackupServer\Backups." Use of this method will correct those pointers to reference 
        /// the path currently specified by the <paramref name="newCatalogPath"/>.</para>
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPCmdletSetSPBackupCatalog"/>
        /// <param name="workingCatalog">An <see cref="SPBackupCatalog"/> object that represents the
        /// backup set to be parsed and path-corrected.</param>
        /// <param name="newCatalogPath">A <c>String</c> containing the new root path that should be used
        /// for modifying the backup directory paths in the <paramref name="workingCatalog"/>.</param>
        /// <remarks>If an exception occurs, it is assumed that the calling method will trap it and
        /// handle it accordingly (including a possible re-throw).</remarks>
        public static void UpdateBackupCatalogPaths(SPBackupCatalog workingCatalog, String newCatalogPath)
        {
            // We need the XML structure for the table of contents file.
            XElement tocFile = workingCatalog.TableOfContents;

            // Perform the modifications
            XElement workingToc = UpdateBackupCatalogPaths(tocFile, newCatalogPath);

            // Assign the updated TOC structure back to the catalog and save it off.
            workingCatalog.TableOfContents = workingToc;
            workingCatalog.SaveTableOfContents();            
        }


        /// <summary>
        /// This method is called when the backup set pointers contained within a <paramref name="tocFile"/>
        /// need to be validated and corrected to make them consistent with the current path held by 
        /// the SPBackupCatalog. This type of correction becomes necessary when a backup catalog that 
        /// was created at one path (for example, "\\BackupServer\Backups") is moved to another path 
        /// (for example, "\\BackupServer\Restores") and referenced. 
        /// <para>In this scenario, folder pointers that are internal to the backup catalog will still 
        /// point to "\\BackupServer\Backups." Use of this method will correct those pointers to reference 
        /// the path currently specified by the <paramref name="newCatalogPath"/>.</para>
        /// </summary>
        /// <seealso cref="SPBackupCatalog"/>
        /// <seealso cref="SPCmdletSetSPBackupCatalog"/>
        /// <param name="tocFile">An <c>XElement</c> object that contains the table of contents XML for the
        /// backup set to be parsed and path-corrected.</param>
        /// <param name="newCatalogPath">A <c>String</c> containing the new root path that should be used
        /// for modifying the backup directory paths in the <paramref name="tocFile"/>.</param>
        /// <remarks>If an exception occurs, it is assumed that the calling method will trap it and
        /// handle it accordingly (including a possible re-throw).</remarks>
        public static XElement UpdateBackupCatalogPaths(XElement tocFile, String newCatalogPath)
        {
            // Create a new working copy of the tocFile
            XElement workingToc = new XElement(tocFile);

            // Work through all history objects - even the ones for things like restores, 
            // configuration-only backups, etc. Everything must be internally consistent.
            foreach (XElement historyObject in workingToc.Elements())
            {
                // Grab the directory name, if it exists, and use that to build a new backup directory
                // specification for use.
                XElement directoryNameElement = historyObject.Element("SPDirectoryName");
                if (directoryNameElement != null)
                {
                    // SharePoint's backup mechanisms include trailing path separators, so we'll do the
                    // same to make sure we're in alignment with OOTB backup is doing.
                    String newBackupDirectory = Path.Combine(newCatalogPath, directoryNameElement.Value);
                    if (!newBackupDirectory.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    {
                        newBackupDirectory += Path.DirectorySeparatorChar;
                    }

                    // Make sure we actually have a BackupDirectory element to modify ...
                    XElement backupDirectoryElement = historyObject.Element("SPBackupDirectory");
                    if (backupDirectoryElement != null)
                    {
                        backupDirectoryElement.SetValue(newBackupDirectory);
                    }
                }
            }

            // Return the resultant (modified) XElement
            return workingToc;
        }


        #endregion UpdateBackupCatalogPaths


        #endregion Methods (Public, Static)


        #region Methods (Private, Static)


        /// <summary>
        /// The purpose of this method is exactly what is stated by the name of the method:
        /// recursive copying (all files and folders) from a source to destination.
        /// </summary>
        private static void RecursiveCopy(String sourceFolder, String destinationFolder)
        {
            // Start by ensuring that the destination folder exists.
            if (!Directory.Exists(destinationFolder))
            { Directory.CreateDirectory(destinationFolder); }

            // Copy every file in the source folder to the destination folder. To do this,
            // we need the full path of each file in the directory
            String[] allFiles = Directory.GetFiles(sourceFolder);
            foreach (String currentFile in allFiles)
            {
                // Extract the filename portion of the fully qualified path and combine that
                // with the destination folder name to build the spec for the file copy.
                String fileName = Path.GetFileName(currentFile);
                String targetFileSpec = Path.Combine(destinationFolder, fileName);
                File.Copy(currentFile, targetFileSpec, true);
            }

            // Now that all files have been copied, we can focus on copying the subdirectories
            // that exist within the source folder.
            String[] allFolders = Directory.GetDirectories(sourceFolder);
            foreach (String currentFolder in allFolders)
            {
                // As with the filename copy, we'll need the name of each individual directory
                // to build the equivalent at the destination folder. 
                //
                // NOTE Although we're using a GetFileName call here, the relative path is what
                // is actually returned.
                String folderName = Path.GetFileName(currentFolder);
                String targetFolderSpec = Path.Combine(destinationFolder, folderName);

                // Make the recursive call to process all contents in the target folder.
                RecursiveCopy(currentFolder, targetFolderSpec);
            }
        }


        #endregion


    }
}
