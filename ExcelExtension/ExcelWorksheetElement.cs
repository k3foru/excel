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
    /// Class for Excel worksheet. 
    /// </summary>
    /// <remarks>Should be visible to COM.</remarks>
    [ComVisible(true)]
    public sealed class ExcelWorksheetElement : ExcelElement
    {
        #region Simple Properties & Methods

        /// <summary>
        /// Gets the class name of this element.
        /// </summary>
        public override string ClassName
        {
            get { return "Excel.Sheet"; }
        }

        /// <summary>
        /// Gets the universal control type of this element.
        /// </summary>
        public override string ControlTypeName
        {
            get
            {
                return ControlType.Table.Name;
            }
        }

        /// <summary>
        /// Gets the name of this element.
        /// </summary>
        public override string Name
        {
            get
            {
                return this.WorksheetInfo.SheetName;
            }
        }

        #endregion

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
                ExcelWorksheetElement otherElement = element as ExcelWorksheetElement;
                if (otherElement != null)
                {
                    return object.Equals(this.WorksheetInfo, otherElement.WorksheetInfo);
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
            return this.WorksheetInfo.GetHashCode();
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
                    this.parent = this.technologyManager.GetExcelElement(this.WindowHandle, null);
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
            int row = 0, column = 0;
            string rowString = condition.GetPropertyValue(PropertyNames.RowIndex) as string;
            string columnString = condition.GetPropertyValue(PropertyNames.ColumnIndex) as string;
            if (int.TryParse(rowString, out row) &&
                int.TryParse(columnString, out column))
            {
                UITechnologyElement cellElement = this.technologyManager.GetExcelElement(this.WindowHandle,
                        new ExcelCellInfo(row, column, this.WorksheetInfo));
                return new UITechnologyElement[] { cellElement }.GetEnumerator();
            }

            return null;
        }

        /// <summary>
        /// Gets or sets the worksheet info.
        /// </summary>
        internal ExcelWorksheetInfo WorksheetInfo { get; private set; }

        /// <summary>
        /// Creates the ExcelWorksheetElement.
        /// </summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <param name="worksheetInfo">The worksheet info.</param>
        /// <param name="manager">The technology manager.</param>
        internal ExcelWorksheetElement(IntPtr windowHandle, ExcelWorksheetInfo worksheetInfo, ExcelTechnologyManager manager)
            : base(windowHandle, manager)
        {
            this.WorksheetInfo = worksheetInfo;
        }

        #endregion
    }
}
