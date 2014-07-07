// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

// Assembly attribute to denote that this assembly has UITest extensions.
[assembly: Microsoft.VisualStudio.TestTools.UITest.Extension.UITestExtensionPackage(
                "ExcelExtensionPackage",
                typeof(Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension.ExcelExtensionPackage))]

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension
{
    using System;
    using Microsoft.VisualStudio.TestTools.UITest.Common;
    using Microsoft.VisualStudio.TestTools.UITest.Extension;
    using Microsoft.VisualStudio.TestTools.UITesting;

    /// <summary>
    /// Entry class for Excel extension package.
    /// </summary>
    internal class ExcelExtensionPackage : UITestExtensionPackage
    {
        /// <summary>
        /// Gets the service object of the specified type.
        /// </summary>
        /// <param name="serviceType">An object that specifies the type of service object to get.</param>
        /// <returns>
        /// A service object of type serviceType or null if there is no service object of type serviceType.
        /// </returns>
        public override object GetService(Type serviceType)
        {
            // Return appropriate service.
            if (serviceType == typeof(UITechnologyManager))
            {
                return technologyManager;
            }
            else if (serviceType == typeof(UITestPropertyProvider))
            {
                return propertyProvider;
            }
            else if (serviceType == typeof(UITestActionFilter))
            {
                return actionFilter;
            }

            return null;
        }

        /// <summary>
        /// Performs application-defined tasks of cleaning up resources.
        /// </summary>
        public override void Dispose()
        {
            // nothing to cleanup
        }

        #region Simple Properties

        /// <summary>
        /// Gets the short description of the package.
        /// </summary>
        public override string PackageDescription
        {
            get { return "Plugin for VSTT Record and Playback support on Excel"; }
        }

        /// <summary>
        /// Gets the name of the package.
        /// </summary>
        public override string PackageName
        {
            get { return "VSTT Excel Plugin"; }
        }

        /// <summary>
        /// Gets the name of the vendor of the package.
        /// </summary>
        public override string PackageVendor
        {
            get { return "Microsoft Corporation"; }
        }

        /// <summary>
        /// Gets the version of the package.
        /// </summary>
        public override Version PackageVersion
        {
            get { return new Version(1, 0); }
        }

        /// <summary>
        /// Gets the version of Visual Studio supported by this package.
        /// </summary>
        public override Version VSVersion
        {
            get { return new Version(10, 0); }
        }

        #endregion

        // Create and cache service.
        private ExcelTechnologyManager technologyManager = new ExcelTechnologyManager();
        private ExcelPropertyProvider propertyProvider = new ExcelPropertyProvider();
        private ExcelActionFilter actionFilter = new ExcelActionFilter();
    }
}
