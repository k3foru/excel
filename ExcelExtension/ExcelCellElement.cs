// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.TestTools.UITest.Extension;
    using Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication;
    using Microsoft.VisualStudio.TestTools.UITesting;

    /// <summary>
    /// Class for Excel cell.
    /// </summary>
    /// <remarks>Should be visible to COM.</remarks>
    [ComVisible(true)]
    public sealed class ExcelCellElement : ExcelElement
    {
        #region Simple Properties & Methods

        /// <summary>
        /// Gets the 0-based position in the parent element's collection.
        /// </summary>
        public override int ChildIndex
        {
            get { return this.CellInfo.RowIndex * this.CellInfo.ColumnIndex; }
        }

        /// <summary>
        /// Gets the class name of this element.
        /// </summary>
        public override string ClassName
        {
            get { return "Excel.Cell"; }
        }

        /// <summary>
        /// Gets the universal control type of this element.
        /// </summary>
        public override string ControlTypeName
        {
            get
            {
                return ControlType.Cell.Name;
            }
        }

        /// <summary>
        /// Gets whether this element is a leaf node (i.e. does not have any children) or not.
        /// </summary>
        public override bool IsLeafNode
        {
            get { return true; }
        }

        /// <summary>
        /// Gets the name of this element.
        /// </summary>
        public override string Name
        {
            get
            {
                return this.CellInfo.ToString();
            }
        }

        /// <summary>
        /// Gets the underlying native technology element (like IAccessible) corresponding this element.
        /// </summary>
        public override object NativeElement
        {
            // Here WindowHandle with CellInfo uniquely identifies underlying Excel control.
            get { return new object[] { this.WindowHandle, this.CellInfo }; }
        }

        /// <summary>
        /// Gets the coordinates of the rectangle that completely encloses this element.
        /// </summary>
        /// <remarks>This is in screen coordinates and never cached.</remarks>
        public override void GetBoundingRectangle(out int left, out int top, out int width, out int height)
        {
            left = top = width = height = -1;

            Utilities.RECT windowRect;
            if (Utilities.GetWindowRect(this.WindowHandle, out windowRect))
            {
                double[] cellRect = ExcelCommunicator.Instance.GetBoundingRectangle(this.CellInfo);

                // Convert the info got from Excel
                //      a) From Point to Pixel.
                //      b) From relative coordinates to screen coordinates.
                left = windowRect.left + Utilities.PointToPixel(cellRect[0], true);
                top = windowRect.top + Utilities.PointToPixel(cellRect[1], false);
                width = Utilities.PointToPixel(cellRect[2], true);
                height = Utilities.PointToPixel(cellRect[3], false);
            }
        }

        /// <summary>
        /// Gets the value for the specified property for this element.
        /// </summary>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns>The value of the property.</returns>
        public override object GetPropertyValue(string propertyName)
        {
            // At the very least, all the properties used in QueryId should be supported here.
            if (string.Equals(PropertyNames.WorksheetName, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.CellInfo.Parent.SheetName;
            }
            else if (string.Equals(PropertyNames.RowIndex, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.CellInfo.RowIndex;
            }
            else if (string.Equals(PropertyNames.ColumnIndex, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.CellInfo.ColumnIndex;
            }

            return base.GetPropertyValue(propertyName);
        }

        #endregion

        /// <summary>
        /// Gets a QueryId that can be used to uniquely identify/find this element.
        /// In some cases, like TreeItem, the QueryIds might contain the entire element hierarchy
        /// but in most cases it will contain only important ancestors of the element.
        /// The technology manager needs to choose which ancestor to capture in the hierarchy
        /// by appropriately setting the QueryId.Ancestor property of each element.
        /// 
        /// The APIs in condition classes like AndCondition.ToString() and AndCondition.Parse()
        /// may be used to convert from this class to string or vice-versa.
        /// </summary>
        public override IQueryElement QueryId
        {
            // For cell, ControlType with RowIndex and ColumnIndex is unique identifier.
            get
            {
                if (this.queryId == null)
                {
                    this.queryId = new QueryElement();
                    this.queryId.Condition = new AndCondition(
                            new PropertyCondition(PropertyNames.ControlType, this.ControlTypeName),
                            new PropertyCondition(PropertyNames.RowIndex, this.CellInfo.RowIndex),
                            new PropertyCondition(PropertyNames.ColumnIndex, this.CellInfo.ColumnIndex));
                    this.queryId.Ancestor = this.Parent;
                }

                return this.queryId;
            }
        }

        #region Override for Object Methods

        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="element">The object to compare with the current object.</param>
        /// <returns>True if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(IUITechnologyElement element)
        {
            if (base.Equals(element))
            {
                ExcelCellElement otherElement = element as ExcelCellElement;
                if (otherElement != null)
                {
                    return this.CellInfo.Equals(otherElement.CellInfo);
                }
            }

            return false;
        }

        /// <summary>
        /// Gets the hash code for this object.
        /// .NET Design Guidelines suggests overridding this too if Equals is overridden. 
        /// </summary>
        /// <returns>The hash code.</returns>
        public override int GetHashCode()
        {
            return this.CellInfo.GetHashCode();
        }

        #endregion

        #region Advance Methods

        /// <summary>
        /// Sets the focus on this element.
        /// </summary>
        public override void SetFocus()
        {
            // Use Excel to set focus (activate) the cell.
            ExcelCommunicator.Instance.SetFocus(this.CellInfo);
        }

        /// <summary>
        /// Scrolls this element into view.
        /// If the technology manager does not support scrolling multiple containers, 
        /// then the outPointX and outPointY should be returned as -1, -1.
        /// </summary>
        /// <param name="pointX">The relative x coordinate of point to make visible.</param>
        /// <param name="pointY">The relative y coordinate of point to make visible.</param>
        /// <param name="outPointX">The relative x coordinate of the point with respect to top most container after scrolling.</param>
        /// <param name="outPointY">The relative y coordinate of the point with respect to top most container after scrolling.</param>
        /// <seealso cref="UITechnologyManagerProperty.ContainerScrollingSupported"/>
        public override void EnsureVisibleByScrolling(int pointX, int pointY, ref int outpointX, ref int outpointY)
        {
            // Use Excel to get the cell into view.
            ExcelCommunicator.Instance.ScrollIntoView(this.CellInfo);
        }

        #endregion

        #region Internals/Privates

        /// <summary>
        /// Gets the parent of this control in this technology.
        /// </summary>
        internal override UITechnologyElement Parent
        {
            get
            {
                if (this.parent == null)
                {
                    this.parent = this.technologyManager.GetExcelElement(this.WindowHandle, this.CellInfo.Parent);
                }

                return this.parent;
            }
        }

        /// <summary>
        /// Gets the children of this control in this technology matching the given condition. 
        /// </summary>
        /// <param name="condition">The condition to match.</param>
        /// <returns>The enumerator for children.</returns>
        internal override System.Collections.IEnumerator GetChildren(AndCondition condition)
        {
            // Cell has no child.
            return null;
        }

        /// <summary>
        /// Gets or sets the cell info.
        /// </summary>
        internal ExcelCellInfo CellInfo { get; private set; }

        /// <summary>
        /// Creates an ExcelCellElement instance.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <param name="cellInfo">The cell info.</param>
        /// <param name="manager">The technology manager.</param>
        internal ExcelCellElement(IntPtr windowHandle, ExcelCellInfo cellInfo, ExcelTechnologyManager manager)
            : base(windowHandle, manager)
        {
            this.CellInfo = cellInfo;
        }

        #endregion
    }
}
