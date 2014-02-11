#region Namespace Imports


// Standard imports
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// For pipe binding
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Backup;
using Microsoft.SharePoint.PowerShell;

// For documentation generation using Gary Lapointe's MAML generator
using Lapointe.PowerShell.MamlGenerator.Attributes;

// Alias the resources
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{
    
    
    /// <summary>
    /// This class provides a number of different ways to get an <see cref="SPBackupCatalog"/> object
    /// created and available for pipe binding purposes.
    /// </summary>
    public class SPBackupCatalogPipeBind : SPCmdletPipeBind<SPBackupCatalog>
    {


        #region Member Declarations


        // Used to store the relevant information for the backup catalog
        private String _catalogPath;


        #endregion Member Declarations


        #region Constructors


        // Default constructor; accepts an SPBackupCata
        public SPBackupCatalogPipeBind(SPBackupCatalog instance)
            : base(instance)
        { }


        // Allows an SPBackupCatalog to be created simply by supplying the path to a backup folder
        public SPBackupCatalogPipeBind(String catalogPath)
        {
            _catalogPath = catalogPath;
        }


        // Takes a farm object and uses its default backup location
        public SPBackupCatalogPipeBind(SPFarm targetFarm)
        {
            SPBackupRestoreConfigurationSettings defaultSettings = SPBackupRestoreConfigurationSettings.CreateSettings(targetFarm);
            _catalogPath = defaultSettings.PreviousBackupLocation;
        }

        
        #endregion Constructors


        #region Overrides (SPCmdletPipeBind)


        /// <summary>
        /// Provides a way to extract critical information that can be used to build another instance
        /// of <see cref="SPBackupCatalog"/> when the <see cref="Read"/> method is called.
        /// </summary>
        /// <param name="instance">The live instance of the <see cref="SPBackupCatalog"/> that will be
        /// examined and captured.</param>
        /// <seealso cref="Read"/>
        /// <seealso cref="SPBackupCatalog"/>
        protected override void Discover(SPBackupCatalog instance)
        {
            _catalogPath = instance.CatalogPath;
        }


        /// <summary>
        /// This method is called when a downstream cmdlet is attempting to get an <see cref="SPBackupCatalog"/>
        /// from the pipe to carry out operations.
        /// </summary>
        /// <returns>An instance of the appropriate <see cref="SPBackupCatalog"/> that was supplied or
        /// referenced at the front end of the pipe.</returns>
        /// <seealso cref="SPBackupCatalog"/>
        public override SPBackupCatalog Read()
        {
            SPBackupCatalog newCatalog = new SPBackupCatalog(_catalogPath);
            if (!newCatalog.IsValidCatalog)
            {
                throw new SPCmdletPipeBindException(String.Format(res.SPBackupCatalogPipeBind_Ex_Read, _catalogPath));
            }
            return newCatalog;
        }


        #endregion Overrides (SPCmdletPipeBind)


    }
}
