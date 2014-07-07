// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelAddin
{
    using System;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication;

    /// <summary>
    /// Implementation of IExcelUITestCommunication which provides information
    /// to the ExcelExtension (loaded in the Coded UI Test process) from the
    /// ExcelAddin (loaded in the Excel process) via .NET Remoting.
    /// </summary>
    internal class UITestCommunicator : MarshalByRefObject, IExcelUITestCommunication
    {
        /// <summary>
        /// Default constructor.
        /// </summary>
        public UITestCommunicator()
        {
            if (ThisAddIn.Instance == null || ThisAddIn.Instance.Application == null)
            {
                throw new InvalidOperationException();
            }

            // Cache the Excel application of this addin.
            this.application = ThisAddIn.Instance.Application;
        }

        #region IExcelUITestCommunication Implementation

        /// <summary>
        /// Gets an Excel UI element at the given screen location. 
        /// </summary>
        /// <param name="x">The x-coordinate of the location.</param>
        /// <param name="y">The y-coordinate of the location.</param>
        /// <returns>The Excel UI element info.</returns>
        public ExcelElementInfo GetElementFromPoint(int x, int y)
        {
            // Use Excel's Object Model to get the required.
            Worksheet ws = this.application.ActiveSheet as Worksheet;
            if (ws != null && this.application.ActiveWindow != null)
            {
                Range cellAtPoint = this.application.ActiveWindow.RangeFromPoint(x, y) as Range;
                if (cellAtPoint != null)
                {
                    return new ExcelCellInfo(cellAtPoint.Row, cellAtPoint.Column, new ExcelWorksheetInfo(ws.Name));
                }
                else
                {
                    return new ExcelWorksheetInfo(ws.Name);
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the Excel UI element current under keyboard focus.
        /// </summary>
        /// <returns>The Excel UI element info.</returns>
        public ExcelElementInfo GetFocussedElement()
        {
            // Use Excel's Object Model to get the required.
            Worksheet ws = this.application.ActiveSheet as Worksheet;
            if (ws != null)
            {
                Range cell = this.application.ActiveCell;
                if (cell != null)
                {
                    return new ExcelCellInfo(cell.Row, cell.Column, new ExcelWorksheetInfo(ws.Name));
                }
                else
                {
                    return new ExcelWorksheetInfo(ws.Name);
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the bounding rectangle of the Excel cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        /// <returns>The bounding rectangle as an array.
        /// The values are relative to the parent window and in Points (instead of Pixels).</returns>
        public double[] GetBoundingRectangle(ExcelCellInfo cellInfo)
        {
            // Use Excel's Object Model to get the required.
            double[] rect = new double[4];
            rect[0] = rect[1] = rect[2] = rect[3] = -1;

            Range cell = GetCell(cellInfo);
            if (cell != null)
            {
                const double xOffset = 25.6; // The constant width of row name column.
                const double yOffset = 36.0; // The constant height of column name row.
                rect[0] = (double)cell.Left + xOffset;
                rect[1] = (double)cell.Top + yOffset;
                rect[2] = (double)cell.Width;
                rect[3] = (double)cell.Height;

                Range visibleRange = this.application.ActiveWindow.VisibleRange;
                if (visibleRange != null)
                {
                    rect[0] -= (double)visibleRange.Left;
                    rect[1] -= (double)visibleRange.Top;
                }
            }

            return rect;
        }

        /// <summary>
        /// Sets focus on a given cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        public void SetFocus(ExcelCellInfo cellInfo)
        {
            // Use Excel's Object Model to get the required.
            Worksheet ws = GetWorksheet(cellInfo.Parent);
            if (ws != null)
            {
                // There could be some other cell under editing. Exit that mode.
                ExitPreviousEditing(ws);

                ws.Activate();
                Range cell = GetCell(cellInfo);
                if (cell != null)
                {
                    cell.Activate();
                }
            }
        }

        /// <summary>
        /// Scrolls a given cell into view.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        public void ScrollIntoView(ExcelCellInfo cellInfo)
        {
            // Use Excel's Object Model to get the required.
            Worksheet ws = GetWorksheet(cellInfo.Parent);
            if (ws != null)
            {
                ws.Activate();
                double[] rect = this.GetBoundingRectangle(cellInfo);
                this.application.ActiveWindow.ScrollIntoView((int)rect[0], (int)rect[1], (int)rect[2], (int)rect[3]);
            }
        }

        /// <summary>
        /// Gets the property of a given cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns>The value of the property.</returns>
        public object GetCellProperty(ExcelCellInfo cellInfo, string propertyName)
        {
            // Use Excel's Object Model to get the required.
            Range cell = GetCell(cellInfo);
            if (cell == null)
            {
                throw new InvalidOperationException();
            }

            switch (propertyName)
            {
                case PropertyNames.Enabled:
                    return true; // TODO - Needed to add support for "locked" cells.

                case PropertyNames.Value:
                    return cell.Value;

                case PropertyNames.Text:
                    return cell.Text;

                case PropertyNames.WidthInChars:
                    return cell.ColumnWidth;

                case PropertyNames.HeightInPoints:
                    return cell.RowHeight;

                case PropertyNames.Formula:
                    return cell.Formula;

                case PropertyNames.WrapText:
                    return cell.WrapText;

                default:
                    throw new NotSupportedException();
            }
        }

        /// <summary>
        /// Sets the property of a given cell.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <param name="propertyValue">The value of the property.</param>
        public void SetCellProperty(ExcelCellInfo cellInfo, string propertyName, object propertyValue)
        {
            // Use Excel's Object Model to get the required.
            Range cell = GetCell(cellInfo);
            if (cell == null)
            {
                throw new InvalidOperationException();
            }

            switch (propertyName)
            {
                case PropertyNames.Value:
                    cell.Value = propertyValue;
                    break;

                case PropertyNames.WidthInChars:
                    cell.ColumnWidth = propertyValue;
                    break;

                case PropertyNames.HeightInPoints:
                    cell.RowHeight = propertyValue;
                    break;

                case PropertyNames.Formula:
                    cell.Formula = propertyValue;
                    break;

                case PropertyNames.WrapText:
                    cell.WrapText = propertyValue;
                    break;

                default:
                    throw new NotSupportedException();
            }
        }

        #endregion

        #region Private Members

        /// <summary>
        /// Gets the Range (cell) from the cell info.
        /// </summary>
        /// <param name="cellInfo">The cell info.</param>
        /// <returns>The Range.</returns>
        private Range GetCell(ExcelCellInfo cellInfo)
        {
            Range cell = null;
            Worksheet ws = GetWorksheet(cellInfo.Parent);
            if (ws != null)
            {
                cell = ws.Cells[cellInfo.RowIndex, cellInfo.ColumnIndex] as Range;
            }

            return cell;
        }

        /// <summary>
        /// Gets the Worksheet from the worksheet info.
        /// </summary>
        /// <param name="sheetInfo">The worksheet info.</param>
        /// <returns>The Worksheet.</returns>
        private Worksheet GetWorksheet(ExcelWorksheetInfo sheetInfo)
        {
            return this.application.Worksheets[sheetInfo.SheetName] as Worksheet;
        }

        /// <summary>
        /// Exit editing mode for any previous cell.
        /// </summary>
        /// <param name="ws">The worksheet.</param>
        private void ExitPreviousEditing(Worksheet ws)
        {
            // TODO - This is hack. Need to find better way to do this.
            // The current logic is to activate another worksheet which fails
            // if there is only one worksheet in the workbook.
            Worksheet otherSheet = ws.Next as Worksheet;
            if (otherSheet == null)
            {
                otherSheet = ws.Previous as Worksheet;
            }

            if (otherSheet != null)
            {
                otherSheet.Activate();
            }
        }

        /// <summary>
        /// The Excel application.
        /// </summary>
        private Application application;

        #endregion
    }
}
