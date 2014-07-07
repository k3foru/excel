// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension
{
    using System;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.Text;
    using Accessibility;

    internal static class Utilities
    {
        /// <summary>
        /// The technology name.
        /// </summary>
        internal const string ExcelTechnologyName = "Excel";

        /// <summary>
        /// The class name of Excel window
        /// </summary>
        internal const string ExcelClassName = "EXCEL7";

        /// <summary>
        /// Checks if the given window is Excel window.
        /// </summary>
        /// <param name="windowHandle">The window handle to check.</param>
        /// <returns>True if the given window is Excel window, false otherwise.</returns>
        internal static bool IsExcelWorksheetWindow(IntPtr windowHandle)
        {
            string className = GetClassName(windowHandle);
            return string.Equals(className, ExcelClassName, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(className, "EXCEL6", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets the window title given a window handle.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <returns>The window title.</returns>
        internal static string GetWindowText(IntPtr windowHandle)
        {
            const int MaxWindowTextLength = 1024;
            StringBuilder textBuilder = new StringBuilder(MaxWindowTextLength);
            if (GetWindowText(windowHandle, textBuilder, textBuilder.Capacity) <= 0)
            {
                Trace.WriteLine(string.Format(CultureInfo.InvariantCulture,
                                         "GetWindowText for handle {0} Writelineed", windowHandle));
            }

            return textBuilder.ToString();
        }

        /// <summary>
        /// Returns the window classname for the given window handle.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <returns>The window classname.</returns>
        internal static string GetClassName(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero)
            {
                return string.Empty;
            }

            const int MaxClassLength = 0x100;
            StringBuilder classNameStringBuilder = new StringBuilder(MaxClassLength);
            if (GetClassName(windowHandle, classNameStringBuilder, classNameStringBuilder.Capacity) <= 0)
            {
                Trace.WriteLine(string.Format(CultureInfo.InvariantCulture,
                                         "GetClassName for handle {0} Writelineed", windowHandle));
            }

            return classNameStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the IAccessible from the given window handle.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <returns>The IAccessible object.</returns>
        internal static IAccessible AccessibleObjectFromWindow(IntPtr windowHandle)
        {
            Guid accessibleGuid = typeof(IAccessible).GUID;
            IAccessible accessible = null;

            if (AccessibleObjectFromWindow(windowHandle, 0, ref accessibleGuid, ref accessible) != 0)
            {
                Trace.WriteLine(string.Format(CultureInfo.InvariantCulture,
                                         "AccessibleObjectFromWindow for handle {0} Writelineed", windowHandle));
            }

            return accessible;
        }

        /// <summary>
        /// Converts dimension in Point to dimension in Pixel.
        /// </summary>
        /// <param name="points">The value in Point.</param>
        /// <param name="xAxis">Whether value is for x-axis or not (i.e. y-axis).</param>
        /// <returns>The value in Pixel.</returns>
        internal static int PointToPixel(double points, bool xAxis)
        {
            const int defaultDpi = 96;
            const int pixelsPerInch = 72;
            if (dpiX == -1)
            {
                dpiX = dpiY = defaultDpi;
                IntPtr deviceHandle = GetDC(IntPtr.Zero);
                if (IntPtr.Zero != deviceHandle)
                {
                    dpiX = GetDeviceCaps(deviceHandle, LOGPIXELSX);
                    dpiY = GetDeviceCaps(deviceHandle, LOGPIXELSY);
                    ReleaseDC(IntPtr.Zero, deviceHandle);
                }
            }

            return (int)(points * (xAxis ? dpiX : dpiY) / pixelsPerInch);
        }

        /// <summary>
        /// Gets the window handle of the window at the given screen location.
        /// </summary>
        /// <param name="pointX">The x-coordinate of the location.</param>
        /// <param name="pointY">The y-coordinate of the location.</param>
        /// <returns>The window handle.</returns>
        /// <remarks>
        /// For some reason, MSDN doc does not match with FxCop results.
        /// MSDN says POINT is 8 byte long whereas FxCop says it is 4 bytes.
        /// Making POINT 4 bytes results in crash indicating probably FxCop
        /// is wrong.  Ignoring the FxCop warning for time being.
        /// </remarks>
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable")]
        [DllImport("user32.dll")]
        internal static extern IntPtr WindowFromPoint(int pointX, int pointY);

        /// <summary>
        /// Gets the bounding rectangle in screen co-ordinates for a window.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <param name="lpRect">Returns the rectangle info.</param>
        /// <returns>Returns true on success</returns>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr windowHandle, out RECT lpRect);

        /// <summary>
        /// The rectangle structure used by GetWindowRect.
        /// </summary>
        /// <remarks>
        /// The System.Drawing.Rectangle cannot be used because member variables do not match -
        /// right & bottom here vs. width & height in System.Drawing.Rectangle.
        /// </remarks>
        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            internal int left;
            internal int top;
            internal int right;
            internal int bottom;
        }

        /// <summary>
        /// Gets the text of the given window.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <param name="windowText">The buffer that will receive the text.</param>
        /// <param name="maxCharCount">The maximum number of characters to copy to the buffer,
        /// including the NULL character.</param>
        /// <returns>If the function succeeds, the return value is the length, in characters,
        /// of the copied string, not including the terminating NULL character.</returns>
        /// <seealso href="http://msdn2.microsoft.com/en-us/library/ms633520.aspx"/>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(
            IntPtr windowHandle, StringBuilder windowText, int maxCharCount);

        /// <summary>
        /// Gets the class name of the given window.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <param name="classNameText">The buffer that will receive the class name.</param>
        /// <param name="maxCharCount">The maximum number of characters to copy to the buffer,
        /// including the NULL character.</param>
        /// <returns>If the function succeeds, the return value is the length, in characters,
        /// of the copied string, not including the terminating NULL character.</returns>
        /// <seealso href="http://msdn2.microsoft.com/en-us/library/ms633582.aspx"/>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr windowHandle, StringBuilder classNameText, int maxCharCount);

        /// <summary>
        /// The AccessibleObjectFromWindow function retrieves the address of the specified interface to the object associated with the given window.
        /// See Documentation at http://msdn2.microsoft.com/en-us/library/ms696137.aspx
        /// </summary>
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr windowHandle, int dwObjectID,
            ref Guid riid, ref IAccessible pAcc);

        /// <summary>
        /// Gets the handle to the DC for the given window.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <returns>The DC handle.</returns>
        [DllImport("user32.dll")]
        private extern static IntPtr GetDC(IntPtr windowHandle);

        /// <summary>
        /// Releases the DC.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <param name="hdc">The DC handle.</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        private extern static int ReleaseDC(IntPtr windowHandle, IntPtr hdc);

        /// <summary>
        /// Gets the device property based on the index.
        /// </summary>
        /// <param name="hdc">The DC handle.</param>
        /// <param name="index">The index of the property to get.</param>
        /// <returns>The value of the property.</returns>
        [DllImport("gdi32.dll")]
        private extern static int GetDeviceCaps(IntPtr hdc, int index);

        // Variables storing DPI settings.
        private static int dpiX = -1;
        private static int dpiY = -1;

        // Indices for the properties used in GetDeviceCaps.
        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;
    }
}
