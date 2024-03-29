﻿// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension
{
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.TestTools.UITest.Extension;
    using Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication;

    /// <summary>
    /// The technology manager for Excel.
    /// </summary>
    /// <remarks>Should be visible to COM.</remarks>
    [ComVisible(true)]
    public sealed class ExcelTechnologyManager : UITechnologyManager
    {
        /// <summary>
        /// Gets the name of the technology supported by this technology manager.
        /// This is the same as the UITechnologyElement.TechnologyName. 
        /// </summary>
        public override string TechnologyName
        {
            get { return Utilities.ExcelTechnologyName; }
        }

        /// <summary>
        /// Gets the support level of this technology manager for the elements(s) in the given window.
        /// The framework uses this function to select the right technology manager for the element. 
        /// </summary>
        /// <param name="windowHandle">The window handle of the element.</param>
        /// <returns>An integer that indicates the level of support provided for the element
        /// by this technology manager. The higher the value the stronger the support.</returns>
        /// <seealso cref="ControlSupport"/>
        public override int GetControlSupportLevel(IntPtr windowHandle)
        {
            // If this is Excel Worksheet window, then we support it!
            return Utilities.IsExcelWorksheetWindow(windowHandle) ?
                   (int)ControlSupport.ControlSpecificSupport : (int)ControlSupport.NoSupport;
        }

        /// <summary>
        /// Gets the focused element i.e. the element that will receive keyboard events at this instance.
        /// </summary>
        /// <param name="handle">The handle of the window which has the focus.
        /// The element which has focus could be the window itself or a
        /// descendant of this window.</param>
        /// <returns>The element that has the focus or null if there is no element with focus.</returns>
        public override IUITechnologyElement GetFocusedElement(IntPtr windowHandle)
        {
            // Simply delegate the call to Excel add-in.
            Debug.Assert(Utilities.IsExcelWorksheetWindow(windowHandle));
            return GetExcelElement(windowHandle, ExcelCommunicator.Instance.GetFocussedElement());
        }

        /// <summary>
        /// Gets the element at the given screen coordinates.
        /// </summary>
        /// <param name="pointX">The x-coordinate of the screen location.</param>
        /// <param name="pointY">The y-coordinate of the screen location.</param>
        /// <returns>The IUITechnologyElement at the screen coordinates specified.</returns>
        public override IUITechnologyElement GetElementFromPoint(int pointX, int pointY)
        {
            // First get the window at that point.
            IntPtr windowHandle = Utilities.WindowFromPoint(pointX, pointY);

            // Then delegate to Excel add-in to get the Excel UI element at that point.
            Debug.Assert(Utilities.IsExcelWorksheetWindow(windowHandle));
            return GetExcelElement(windowHandle, ExcelCommunicator.Instance.GetElementFromPoint(pointX, pointY));
        }

        /// <summary>
        /// Gets the element from the given window handle.
        /// </summary>
        /// <param name="handle">The window handle.</param>
        /// <returns>The IUITechnologyElement from the window handle.</returns>
        public override IUITechnologyElement GetElementFromWindowHandle(IntPtr windowHandle)
        {
            // Get the Excel worksheet from window handle.
            Debug.Assert(Utilities.IsExcelWorksheetWindow(windowHandle));
            return GetExcelElement(windowHandle, null);
        }

        /// <summary>
        /// Gets the element from the given native (underlying) technology element.
        /// </summary>
        /// <param name="nativeElement">The native technology element (like IAccessible).</param>
        /// <returns>The IUITechnologyElement from the native element.</returns>
        /// <seealso cref="UITechnologyElement.NativeElement"/>
        public override IUITechnologyElement GetElementFromNativeElement(object nativeElement)
        {
            object[] parts = nativeElement as object[];
            if (parts != null && parts.Length == 2 && parts[0] is IntPtr && parts[1] is ExcelElementInfo)
            {
                // Get the cell or worksheet as appropriate.
                IntPtr windowHandle = (IntPtr)parts[0];
                ExcelElementInfo elementInfo = (ExcelElementInfo)parts[1];
                return GetExcelElement(windowHandle, elementInfo);
            }
            else if (nativeElement is IntPtr)
            {
                // For window handle, get the Excel worksheet.
                return GetElementFromWindowHandle((IntPtr)nativeElement);
            }

            return null;
        }

        /// <summary>
        /// Converts the given element of another technology to new element of this technology manager.
        /// This is used for operations like switching between hosted and hosting technologies.
        /// </summary>
        /// <param name="elementToConvert">The element to convert.</param>
        /// <param name="supportLevel">The level of support provided for the
        /// converted element by this technology manager.</param>
        /// <returns>The new converted element in this technology or null if no conversion is possible.</returns>
        /// <seealso cref="GetControlSupportLevel"/>
        /// <seealso cref="ControlSupport"/>
        public override IUITechnologyElement ConvertToThisTechnology(IUITechnologyElement elementToConvert, out int supportLevel)
        {
            supportLevel = (int)ControlSupport.NoSupport;
            if (elementToConvert is ExcelElement)
            {
                // If already an Excel UI element, no need to convert.
                supportLevel = (int)ControlSupport.ControlSpecificSupport;
                return elementToConvert;
            }
            else
            {
                // If this is an Excel worksheet window, convert appropriate.
                IntPtr windowHandle = elementToConvert.WindowHandle;
                if (Utilities.IsExcelWorksheetWindow(windowHandle))
                {
                    supportLevel = (int)ControlSupport.ControlSpecificSupport;
                    return GetExcelElement(windowHandle, null);
                }
            }

            // Return null for other cases.
            return null;
        }

        /// <summary>
        /// Parses the query element string and returns the parsedQueryIdCookie to be
        /// used later during Search() or MatchElement() or GetChildren() call for
        /// either searching or matching or getting children that has the same query string.
        /// </summary>
        /// <param name="queryElement">The query element string to parse.</param>
        /// <param name="parsedQueryIdCookie">The cookie of the parsed QueryId to be used later.</param>
        /// <returns>The remaining part of query element string that is not supported
        /// by this technology manager. The framework may or may not support the remaining part.</returns>
        public override string ParseQueryId(string queryElement, out object parsedQueryIdCookie)
        {
            // Use the AndCondition.Parse() API to get the condition object
            // from string. Here, all the properties possible in query id are
            // supported by this technology manager itself.
            IQueryCondition condition = null;
            try
            {
                condition = AndCondition.Parse(queryElement);
            }
            catch (ArgumentException)
            {
            }

            if (condition == null)
            {
                // Implies parse failed. This should not be the case.
                Debug.Fail("ParseQueryId failed");
                parsedQueryIdCookie = null;
                return queryElement;
            }
            else
            {
                // Store the condition as the cookie to be used later.
                parsedQueryIdCookie = condition;
                return string.Empty;
            }
        }

        /// <summary>
        /// Matches the element against the parsedQueryIdCookie condition.
        /// </summary>
        /// <param name="element">The element to match against the conditions.</param>
        /// <param name="parsedQueryIdCookie">The cookie of previously parsed QueryId.</param>
        /// <param name="useEngine">
        /// This is set to true by the technology manager if it wants to use
        /// the framework for matching the complete or part of query element.</param>
        /// <returns>True if the element matches the condition, false otherwise.</returns>
        /// <remarks>
        /// This is an optional method and if the technology manager does not support
        /// this method it should throw System.NotSupportedException exception. If the Search()
        /// is not supported then the framework uses GetChildren() API to do breadth-first
        /// traversal and for each element uses MatchElement() API to match & find.
        /// 
        /// Note that a technology has to support either this or Search.
        /// </remarks>
        /// <seealso cref="ParseQueryId"/>
        /// <seealso cref="Search"/>
        public override bool MatchElement(IUITechnologyElement element, object parsedQueryIdCookie, out bool useEngine)
        {
            // Get the condition out of the cookie as set by ParseQueryId() API.
            IQueryCondition condition = parsedQueryIdCookie as AndCondition;
            if (condition != null)
            {
                // Use the Match API to do the matching. Note that this API
                // will call into ExcelElement.GetPropertyValue() method to
                // get the value of the properties for matching. 
                useEngine = false;
                return condition.Match(element);
            }
            else
            {
                useEngine = true;
                return false;
            }
        }

        #region Navigation Methods

        /// <summary>
        /// Gets the parent of the given element in the user interface hierarchy.
        /// </summary>
        /// <param name="element">The element whose parent is needed.</param>
        /// <returns>The parent element or null if the element passed is the
        /// root element in this technology.</returns>
        public override IUITechnologyElement GetParent(IUITechnologyElement element)
        {
            ExcelElement excelElement = element as ExcelElement;
            if (excelElement != null)
            {
                return excelElement.Parent;
            }

            return null;
        }

        /// <summary>
        /// Gets the enumerator for children of the given IUITechnologyElement.
        /// </summary>
        /// <param name="element">The IUITechnologyElement whose child enumerator is needed.</param>
        /// <param name="parsedQueryIdCookie">The cookie of previously parsed QueryId to filter matching children.</param>
        /// <returns>The enumerator for children.</returns>
        /// <seealso cref="ParseQueryId"/>
        public override System.Collections.IEnumerator GetChildren(IUITechnologyElement element, object parsedQueryIdCookie)
        {
            ExcelElement excelElement = element as ExcelElement;
            AndCondition condition = parsedQueryIdCookie as AndCondition;
            if (excelElement != null)
            {
                if (condition != null)
                {
                    return excelElement.GetChildren(condition);
                }
                else
                {
                    return new UITechnologyElement[] { }.GetEnumerator();
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the next sibling of the given element in the user interface hierarchy.
        /// </summary>
        /// <param name="element">The element whose next sibling is needed.</param>
        /// <returns>The next sibling or null if none is present.</returns>
        public override IUITechnologyElement GetNextSibling(IUITechnologyElement element)
        {
            // TODO - Sibling navigation is required to get the arrow keys of Spy control working. 
            return null;
        }

        /// <summary>
        /// Gets the previous sibling of the given element in the user interface hierarchy.
        /// </summary>
        /// <param name="element">The element whose previous sibling is needed.</param>
        /// <returns>The previous sibling or null if none is present.</returns>
        public override IUITechnologyElement GetPreviousSibling(IUITechnologyElement element)
        {
            // TODO - Sibling navigation is required to get the arrow keys of Spy control working. 
            return null;
        }

        #endregion

        #region Search Methods - Not required for this sample.

        /// <summary>
        /// Searches for an element matching the given query element. If the underlying
        /// UI Technology has rich APIs to search/navigate the UI hierarchy then
        /// implementing this method could improve the playback performance significantly.
        /// </summary>
        /// <param name="parsedQueryIdCookie">The cookie of previously parsed QueryId.</param>
        /// <param name="parentElement">The parent element under which to search.</param>
        /// <param name="maxDepth">The maximum tree depth to search.</param>
        /// <returns>An array of matched elements or null if none matched.</returns>
        /// <remarks>
        /// This is an optional method and if the technology manager does not support
        /// this method it should throw System.NotSupportedException exception. If this
        /// is not supported then the framework uses GetChildren() API to do breadth-first
        /// traversal and for each element uses MatchElement() API to match & find.
        /// 
        /// Note that a technology has to support either this or MatchElement.
        /// </remarks>
        /// <seealso cref="ParseQueryId"/>
        /// <seealso cref="MatchElement"/>
        /// <seealso cref="UITechnologyManagerProperty.SearchSupported"/>
        public override object[] Search(object parsedQueryIdCookie, IUITechnologyElement parentElement, int maxDepth)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Gets the information about the most recent invocation of the technology manager.
        /// </summary>
        /// <returns>Information about the most recent invocation of the technology manager.</returns>
        public override ILastInvocationInfo GetLastInvocationInfo()
        {
            return null;
        }

        /// <summary>
        /// Cancels any wait or search operation being performed by this technology manager
        /// because of call to WaitForReady or Search methods.
        /// </summary>
        /// <remarks>
        /// This call is made on another thread as both the WaitForReady and Search methods are blocking.
        /// </remarks>
        public override void CancelStep()
        {
            // no op
        }

        #endregion

        #region Add/Remove Event Methods - Not required for this sample.

        /// <summary>
        /// Adds an event handler.
        /// </summary>
        /// <param name="element">The element and its descendants for which this event should be fired.</param>
        /// <param name="eventType">The type of event to listen to.</param>
        /// <param name="eventSink">The event sink which should be notified
        /// when the event occurs.</param>
        /// <returns>True if the eventType is supported and add is successful, false otherwise.</returns>
        public override bool AddEventHandler(IUITechnologyElement element, UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        /// <summary>
        /// Removes an event handler.
        /// </summary>
        /// <param name="element">The element and its descendants for which this event should be removed.</param>
        /// <param name="eventType">The type of event to listen to.</param>
        /// <param name="eventSink">The event sink interface that was registered.</param>
        /// <returns>True if the eventType is supported and remove is successful, false otherwise.</returns>
        public override bool RemoveEventHandler(IUITechnologyElement element, UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        /// <summary>
        /// Adds a global sink to notifiy actions raised by the plugin
        /// </summary>
        /// <param name="eventType">The type of event to listen to.</param>
        /// <param name="eventSink">Sink used for notification</param>
        /// <returns>True if successful, otherwise false.</returns>
        public override bool AddGlobalEventHandler(UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        /// <summary>
        /// Removes a global sink to notifiy actions raised by the plugin
        /// </summary>
        /// <param name="eventType">The type of event to remove.</param>
        /// <param name="eventSink">Sink used for notification</param>
        /// <returns>True if successful, otherwise false.</returns>
        public override bool RemoveGlobalEventHandler(UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        /// <summary>
        /// Gets a synchronization waiter for given UITestEventType on this element.
        /// </summary>
        /// <param name="element">The element to get synchronization waiter for.</param>
        /// <param name="eventType">The event for which synchronization waiter is required.</param>
        /// <returns>
        /// The synchronization waiter for specified event or null if event/waiter is not supported.
        /// </returns>
        public override IUISynchronizationWaiter GetSynchronizationWaiter(IUITechnologyElement element, UITestEventType eventType)
        {
            return null;
        }

        /// <summary>
        /// Processes the process mouse enter event for the window.
        /// </summary>
        /// <param name="handle">The window handle.</param>
        public override void ProcessMouseEnter(IntPtr handle)
        {
            // no op
        }

        #endregion

        #region Initialize/Cleanup Methods - Not required for this sample

        /// <summary>
        /// Performs any initialization required by this technology manager for starting a session.
        /// </summary>
        /// <param name="recordingSession">True if this is recording session, false otherwise like for playback session.</param>
        public override void StartSession(bool recordingSession)
        {
            // no op
        }

        /// <summary>
        /// Performs any cleanup required by this technology manager for stopping the current session.
        /// </summary>
        public override void StopSession()
        {
            // no op
        }

        #endregion

        #region Internal/Private

        /// <summary>
        /// Creates an appropriate Excel UI element. 
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <param name="elementInfo">The info on the element.</param>
        /// <returns>The Excel UI element.</returns>
        internal UITechnologyElement GetExcelElement(IntPtr windowHandle, ExcelElementInfo elementInfo)
        {
            if (elementInfo is ExcelCellInfo)
            {
                return new ExcelCellElement(windowHandle, elementInfo as ExcelCellInfo, this);
            }
            else if (elementInfo is ExcelWorksheetInfo)
            {
                return new ExcelWorksheetElement(windowHandle, elementInfo as ExcelWorksheetInfo, this);
            }
            else
            {
                return new ExcelElement(windowHandle, this);
            }
        }

        #endregion
    }
}
