// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication
{
    using System;
    using System.Globalization;

    /// <summary>
    /// This interface is used by the ExcelExtension (loaded in the Coded UI Test process)
    /// to communicate with the ExcelAddin (loaded in the Excel process) via .NET Remoting.
    /// </summary>
    public interface IExcelUITestCommunication
    {
        /// <summary>
        /// Gets an Excel UI element at the given screen location. 
        /// </summary>
        /// <param name="x">The x-coordinate of the location.</param>
        /// <param name="y">The y-coordinate of the location.</param>
        /// <returns>The Excel UI element info.</returns>
        ExcelElementInfo GetElementFromPoint(int x, int y);

        /// <summary>
        /// Gets the Excel UI element current under keyboard focus.
        /// </summary>
        /// <returns>The Excel UI element info.</returns>
        ExcelElementInfo GetFocussedElement();

        /// <summary>
        /// Gets the bounding rectangle of the Excel cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        /// <returns>The bounding rectangle as an array.
        /// The values are relative to the parent window and in Points (instead of Pixels).</returns>
        double[] GetBoundingRectangle(ExcelCellInfo cellInfo);

        /// <summary>
        /// Sets focus on a given cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        void SetFocus(ExcelCellInfo cellInfo);

        /// <summary>
        /// Scrolls a given cell into view.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        void ScrollIntoView(ExcelCellInfo cellInfo);

        /// <summary>
        /// Gets the property of a given cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns>The value of the property.</returns>
        object GetCellProperty(ExcelCellInfo cellInfo, string propertyName);

        /// <summary>
        /// Sets the property of a given cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <param name="propertyValue">The value of the property.</param>
        void SetCellProperty(ExcelCellInfo cellInfo, string propertyName, object propertyValue);
    }

    /// <summary>
    /// Abstract base class for all Excel UI element info.
    /// </summary>
    [Serializable]
    public abstract class ExcelElementInfo
    {
    }

    /// <summary>
    /// Class for Excel worksheet info.
    /// </summary>
    [Serializable]
    public class ExcelWorksheetInfo : ExcelElementInfo
    {
        /// <summary>
        /// Creates a new ExcelWorksheetInfo with the given worksheet name.
        /// </summary>
        /// <param name="sheetName">The name of the worksheet.</param>
        public ExcelWorksheetInfo(string sheetName)
        {
            if (sheetName == null) throw new ArgumentNullException("sheetName");
            SheetName = sheetName;
        }

        /// <summary>
        /// Gets or sets the name of the worksheet.
        /// </summary>
        public string SheetName { get; private set; }

        #region Object Overrides

        // Helpful in debugging.
        public override string ToString()
        {
            return SheetName;
        }

        // Needed to find out if two objects are same or not.
        public override bool Equals(object obj)
        {
            ExcelWorksheetInfo other = obj as ExcelWorksheetInfo;
            if (other != null)
            {
                return string.Equals(SheetName, other.SheetName, StringComparison.Ordinal);
            }

            return false;
        }

        // Good practice to override this when overriding Equals.
        public override int GetHashCode()
        {
            return SheetName.GetHashCode();
        }

        #endregion
    }

    /// <summary>
    /// Class for Excel cell info.
    /// </summary>
    [Serializable]
    public class ExcelCellInfo : ExcelElementInfo
    {
        /// <summary>
        /// Creates a new ExcelCellInfo with the given info.
        /// </summary>
        /// <param name="rowIndex">The row index.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="parent">The parent worksheet.</param>
        public ExcelCellInfo(int rowIndex, int columnIndex, ExcelWorksheetInfo parent)
        {
            if (parent == null) throw new ArgumentNullException("parent");

            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            Parent = parent;
        }

        /// <summary>
        /// Gets or sets the row index of the cell.
        /// </summary>
        public int RowIndex { get; private set; }

        /// <summary>
        /// Gets or sets the column index of the cell.
        /// </summary>
        public int ColumnIndex { get; private set; }

        /// <summary>
        /// Gets or sets the parent worksheet of the cell.
        /// </summary>
        public ExcelWorksheetInfo Parent { get; private set; }

        #region Object Overrides

        // Helpful in debugging.
        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "{0}[{1}, {2}]", Parent, RowIndex, ColumnIndex);
        }

        // Needed to find out if two objects are same or not.
        public override bool Equals(object obj)
        {
            ExcelCellInfo other = obj as ExcelCellInfo;
            if (other != null)
            {
                return RowIndex == other.RowIndex && ColumnIndex == other.ColumnIndex && object.Equals(Parent, other.Parent);
            }

            return false;
        }

        // Good practice to override this when overriding Equals.
        public override int GetHashCode()
        {
            return RowIndex.GetHashCode() ^ ColumnIndex.GetHashCode() ^ Parent.GetHashCode();
        }

        #endregion
    }

    /// <summary>
    /// Names of various properties of an Excel cell.
    /// </summary>
    public static class PropertyNames
    {
        public const string ControlType = "ControlType";
        public const string ClassName = "ClassName";
        public const string Name = "Name";
        public const string WorksheetName = "WorksheetName";
        public const string RowIndex = "RowIndex";
        public const string ColumnIndex = "ColumnIndex";

        public const string Enabled = "Enabled";
        public const string Value = "Value";
        public const string Text = "Text";
        public const string WidthInChars = "WidthInChars";
        public const string HeightInPoints = "HeightInPoints";
        public const string Formula = "Formula";
        public const string WrapText = "WrapText";
    }
}
