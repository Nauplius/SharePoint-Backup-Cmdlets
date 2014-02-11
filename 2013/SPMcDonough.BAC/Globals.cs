#region Namespace Imports


// Standard namespace imports
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// For SharePoint and PowerShell references
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;

// Alias the resources
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{


    /// <summary>
    /// This enumeration is used primarily by the <see cref="SPCmdletExportSPBackupCatalog"/> 
    /// type to simplify the process of identifying how an export/archive operation is
    /// supposed to take place
    /// </summary>
    /// <seealso cref="SPCmdletExportSPBackupCatalog"/>
    internal enum BackupSetExportMode
    {
        Undefined = 0,
        ExportAll = 1,
        IncludeMode = 2,
        ExcludeMode = 3,
    }


    /// <summary>
    /// This class is simply used in cases where a PowerShell error needs to be tracked.
    /// Properties can be set and then used when throwing a terminating exception or writing
    /// error info out.
    /// </summary>
    internal class PowerShellError
    {
        public Object Object { get; set; }
        public ErrorCategory Category { get; set; }
    }


    /// <summary>
    /// This static class houses constants and related types that are used throughout the
    /// solution.
    /// </summary>
    internal static class Globals
    {


        #region Constants


        // Backup set-related constants
        public const String TOC_FILENAME = "spbrtoc.xml";   // houses table of contents for backup set
        
        // For use in generated exports
        public const String EXPORT_FILE_OR_PATH_TEMPLATE = "BackupCatalogExport_{0}";
        public const String EXPORT_DATE_TIME_TEMPLATE = "{0}_{1}_{2}-{3}_{4}_{5}";

        // For use in writing out compressed archives
        public const String EXPORT_ARCHIVE_EXTENSION = ".zip";

        // This basic XML string is used as the basis for which an export catalog is built for
        // the SPCmdExportSPBackupCatalog class.
        public const String BACKUP_CATALOG_BASE_XML = "<SPBackupRestoreHistory/>";

        
        #endregion Constants


        #region Methods

        
        /// <summary>
        /// This method is responsible for taking the supplied date and time and creating
        /// String version of it that can be used in file names and paths.
        /// </summary>
        /// <param name="selectedDateTime">A <c>DateTime</c> that will be used to create the
        /// output string.</param>
        /// <returns>A String that contains a file and path-usable form of the date/time
        /// supplied as <paramref name="selectedDateTime"/></returns>
        public static String BuildFileAndPathCompatibleDateTime(DateTime selectedDateTime)
        {
            String filePathDateTime = String.Format(
                EXPORT_DATE_TIME_TEMPLATE,
                selectedDateTime.Year,
                selectedDateTime.Month,
                selectedDateTime.Day,
                selectedDateTime.Hour,
                selectedDateTime.Minute,
                selectedDateTime.Second);
            return filePathDateTime;
        }


        /// <summary>
        /// The purpose of this method is to take a value (in bytes) and convert it to a nicely
        /// formatted string that is sized appropriately.
        /// </summary>
        /// <param name="byteCount">A nullable Int64 that represents the number of bytes for conversion.</param>
        /// <returns>A String that can be directly inserted into a formatted string to intelligently 
        /// represent the number of bytes passed in</returns>
        public static String IntelligentBytesFormat(Int64? byteCount)
        {
            String formattedBytes = res.SPGlobals_NullValue;
            if (byteCount.HasValue)
            {
                Double byteWork = Convert.ToDouble(byteCount);
                if (byteCount < 1024)
                {
                    formattedBytes = String.Format("{0} bytes", byteWork);
                }
                else if (byteCount < (1024 * 1024))
                {
                    formattedBytes = String.Format("{0:0.00} KB", (byteWork / 1024));
                }
                else if (byteCount < (1024 * 1024 * 1024))
                {
                    formattedBytes = String.Format("{0:0.00} MB", (byteWork / (1024 * 1024)));
                }
                else
                {
                    formattedBytes = String.Format("{0:0.00} GB", (byteWork / (1024 * 1024 * 1024)));
                }
            }

            return formattedBytes;
        }


        /// <summary>
        /// This method takes a nullable DateTime and renders it as something a little more readable for
        /// the average human being.
        /// </summary>
        /// <param name="targetDateTime">A nullable DateTime containing the value that will be formatted.</param>
        /// <returns>A String representation of the date and time that can be inserted into a report
        /// or e-mail.</returns>
        public static String IntelligentDateTimeFormat(DateTime? targetDateTime)
        {
            String formattedDateTime = res.SPGlobals_NullValue;            
            TimeZone localZone = TimeZone.CurrentTimeZone;
            if (targetDateTime.HasValue)
            {
                DateTime localDateTime = targetDateTime.Value.ToLocalTime();
                formattedDateTime = localDateTime.ToShortTimeString() + " on " + localDateTime.ToLongDateString();
                if (localZone.IsDaylightSavingTime(localDateTime))
                {
                    formattedDateTime += String.Format(" ({0})", localZone.DaylightName);
                }
                else
                {
                    formattedDateTime += String.Format(" ({0})", localZone.StandardName);
                }
            }

            return formattedDateTime;
        }


        /// <summary>
        /// This method takes a nullable Integer and renders it as something a little more readable for
        /// the average human being.
        /// </summary>
        /// <param name="targetInt">A nullable Int32 containing the value that will be formatted.</param>
        /// <returns>A String representation of the integer that can be inserted into a report
        /// or e-mail.</returns>
        public static String IntelligentIntegerFormat(Int32? targetInt)
        {
            String formattedInt = res.SPGlobals_NullValue;
            if (targetInt.HasValue)
            {
                formattedInt = targetInt.Value.ToString();
            }

            return formattedInt;
        }
        
        
        /// <summary>
        /// This method takes a nullable TimeSpan and renders it as something a bit more humanly readable
        /// than just some colon-separated values.
        /// </summary>
        /// <param name="amountOfTime">A nullable TimeSpan containing the time that will be formatted.</param>
        /// <returns>A String representation of the time that can be inserted into a report
        /// or e-mail.</returns>
        public static String IntelligentTimeFormat(TimeSpan? amountOfTime)
        {
            // Pull component values out of time.
            Int32 numberOfDays = amountOfTime.Value.Days;
            Int32 numberOfHours = amountOfTime.Value.Hours;
            Int32 numberOfMins = amountOfTime.Value.Minutes;
            Int32 numberOfSeconds = amountOfTime.Value.Seconds;
            
            // Begin building out the time string.
            String formattedTime = String.Empty;
            if (amountOfTime.HasValue)
            {
                if (numberOfDays > 0)
                {
                    formattedTime += (formattedTime.Length == 0 ? String.Format("{0} days", numberOfDays) : String.Format(", {0} days", numberOfDays));
                }
                if (numberOfHours > 0)
                {
                    formattedTime += (formattedTime.Length == 0 ? String.Format("{0} hours", numberOfHours) : String.Format(", {0} hours", numberOfHours));
                }
                if (numberOfMins > 0)
                {
                    formattedTime += (formattedTime.Length == 0 ? String.Format("{0} minutes", numberOfMins) : String.Format(", {0} minutes", numberOfMins));
                }
                if (numberOfSeconds > 0)
                {
                    formattedTime += (formattedTime.Length == 0 ? String.Format("{0} seconds", numberOfSeconds) : String.Format(", {0} seconds", numberOfSeconds));
                }
                if (formattedTime.Contains(","))
                {
                    Int32 lastCommaIdx = formattedTime.LastIndexOf(",");
                    formattedTime = formattedTime.Substring(0, lastCommaIdx) + " and" + formattedTime.Substring(lastCommaIdx + 1);
                }
            }
            else
            {
                formattedTime = res.SPGlobals_NullValue;
            }

            return formattedTime;
        }


        #endregion Methods


    }
}
