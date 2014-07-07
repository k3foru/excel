// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelAddin
{
    using System;
    using System.Runtime.Remoting;
    using System.Runtime.Remoting.Channels;
    using System.Runtime.Remoting.Channels.Ipc;
    using Microsoft.Office.Tools.Excel;
    using Microsoft.Office.Tools.Excel.Extensions;
    using Excel = Microsoft.Office.Interop.Excel;
    using Office = Microsoft.Office.Core;

    public partial class ThisAddIn
    {
        /// <summary>
        /// Singleton instance to this add-in.
        /// </summary>
        internal static ThisAddIn Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Register for .NET Remoting on startup of this add-in.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event arguments.</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Instance = this;
            channel = new IpcChannel("ExcelUITest");
            ChannelServices.RegisterChannel(channel, false);
            RemotingConfiguration.RegisterWellKnownServiceType(typeof(UITestCommunicator),
                "ExcelUITest", WellKnownObjectMode.Singleton);
        }

        /// <summary>
        /// Unregister for .NET Remoting on shutdown of this add-in.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event arguments.</param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (channel != null)
            {
                ChannelServices.UnregisterChannel(channel);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        /// <summary>
        /// The channel for .NET Remoting calls.
        /// </summary>
        private IChannel channel;
    }
}
