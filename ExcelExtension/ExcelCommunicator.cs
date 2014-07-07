// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension
{
    using System;
    using Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication;

    /// <summary>
    /// Static class to manage the Excel communication interface.
    /// </summary>
    internal static class ExcelCommunicator
    {
        /// <summary>
        /// Singleton interface used to communicate with the Excel via .NET Remoting.
        /// </summary>
        internal static IExcelUITestCommunication Instance
        {
            get
            {
                if (excelCommunicator == null)
                {
                    excelCommunicator = (IExcelUITestCommunication)Activator.GetObject(typeof(IExcelUITestCommunication), "ipc://ExcelUITest/ExcelUITest");
                }

                return excelCommunicator;
            }
        }

        /// <summary>
        /// Singleton interface used to communicate with the Excel via .NET Remoting.
        /// </summary>
        private static IExcelUITestCommunication excelCommunicator;
    }
}
