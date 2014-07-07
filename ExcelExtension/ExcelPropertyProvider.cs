// Sample code developed by gautamg@microsoft.com
// Copyright (c) Microsoft Corporation. All rights reserved.

namespace Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension
{
    using System;
    using System.Collections.Generic;
    using Microsoft.VisualStudio.TestTools.UITest.Extension;
    using Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication;
    using Microsoft.VisualStudio.TestTools.UITesting;

    /// <summary>
    /// Property Provider for the Excel cell.
    /// </summary>
    internal class ExcelPropertyProvider : UITestPropertyProvider
    {
        /// <summary>
        /// Gets the support level of the provider for the specified control.
        /// </summary>
        /// <param name="uiTestControl">The control for which the support level is required.</param>
        /// <returns>The support level for the control.</returns>
        /// <seealso cref="Microsoft.VisualStudio.TestTools.UITest.Extension.ControlSupport"/>
        public override int GetControlSupportLevel(UITestControl uiTestControl)
        {
            // Return high value if Excel cell.
            if (string.Equals(uiTestControl.TechnologyName, Utilities.ExcelTechnologyName, StringComparison.OrdinalIgnoreCase) &&
                uiTestControl.ControlType == ControlType.Cell)
            {
                return (int)ControlSupport.ControlSpecificSupport;
            }

            return (int)ControlSupport.NoSupport;
        }

        /// <summary>
        /// Gets all properties supported by this provider for the specified control.
        /// </summary>
        /// <param name="uiTestControl">The control whose properties are required.</param>
        /// <returns>The collection of supported properties.</returns>
        public override ICollection<string> GetPropertyNames(UITestControl uiTestControl)
        {
            // Return name from pre-built dictionary.
            return propertiesMap.Keys;
        }

        /// <summary>
        /// Gets the property descriptor for the given property of the control.
        /// </summary>
        /// <param name="uiTestControl">The control for which descriptor is required.</param>
        /// <param name="propertyName">The name of the property for which descriptor is required.</param>
        /// <returns>The property descriptor of the property.</returns>
        public override UITestPropertyDescriptor GetPropertyDescriptor(UITestControl uiTestControl, string propertyName)
        {
            // Return descriptor from pre-built dictionary.
            if (propertiesMap.ContainsKey(propertyName))
            {
                return propertiesMap[propertyName];
            }

            return null;
        }

        /// <summary>
        /// Gets the property value of the control.
        /// This method is called for all the properties defined in UITestControl.PropertyNames
        /// to let the provider override the values. The provider should either support those
        /// properties properly or throw System.NotSupportedException.
        /// </summary>
        /// <param name="uiTestControl">The control for which to get the value.</param>
        /// <param name="propertyName">The name of the property to get.</param>
        /// <returns>The value of the property.</returns>
        /// <exception cref="System.NotSupportedException">Property with the given name is not supported.</exception>
        public override object GetPropertyValue(UITestControl uiTestControl, string propertyName)
        {
            // Simply delegate the call to Excel add-in.
            ExcelCellInfo cellInfo = GetCellInfo(uiTestControl);
            return ExcelCommunicator.Instance.GetCellProperty(cellInfo, propertyName);
        }

        /// <summary>
        /// Sets the property value of the control.
        /// This method is called for all the settable properties defined in UITestControl.PropertyNames
        /// to let the provider override the values. The provider should either support setting of those
        /// properties properly or throw System.NotSupportedException.
        /// </summary>
        /// <param name="uiTestControl">The control for which to set the value.</param>
        /// <param name="propertyName">The name of the property to set.</param>
        /// <param name="value">The value of the property.</param>
        /// <exception cref="System.NotSupportedException">Property with the given name is not supported.</exception>
        public override void SetPropertyValue(UITestControl uiTestControl, string propertyName, object propertyValue)
        {
            // Just to match real user behavior closely, set Value & Formula
            // properties by actual typing of the text in the cell.
            if (string.Equals(propertyName, PropertyNames.Value, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(propertyName, PropertyNames.Formula, StringComparison.OrdinalIgnoreCase))
            {
                Keyboard.SendKeys(uiTestControl, propertyValue as string);
            }

            // Simply delegate the call to Excel add-in for the rest.
            ExcelCellInfo cellInfo = GetCellInfo(uiTestControl);
            ExcelCommunicator.Instance.SetCellProperty(cellInfo, propertyName, propertyValue);
        }

        #region Code Generation Customization Methods - Not Implemented in this sample

        /// <summary>
        /// Gets the specialized class to use for the specified control.
        /// </summary>
        /// <param name="uiTestControl">The control for which a specialized class is to needed.</param>
        /// <returns>Type object of the specialized class.</returns>
        public override Type GetSpecializedClass(UITestControl uiTestControl)
        {
            return null;
        }

        /// <summary>
        /// Gets the search properties present by default in the specified specialized class.
        /// </summary>
        /// <param name="specializedClassName">System.Type object of the specialized class.</param>
        /// <returns>Properties already present in the specialized class.</returns>
        public override string[] GetPredefinedSearchProperties(Type specializedClassType)
        {
            return null;
        }

        /// <summary>
        /// Gets the property name for the specified action.
        /// For example, to generate code for a SetValueAction on a Edit like: -
        ///         myEdit.Text = "abc";
        /// this method should return "Text" as property name. Otherwise, if this
        /// returns null, then the generated code will look like -
        ///         myEdit.SetProperty("Value", "abc");
        /// </summary>
        /// <param name="uiTestControl">The control on which the action was performed.</param>
        /// <param name="action">The action for which the property is required.</param>
        /// <returns>The writable property name for the action or null if no property exists.</returns>
        public override string GetPropertyForAction(UITestControl uiTestControl, Common.UITestAction action)
        {
            return null;
        }

        /// <summary>
        /// Gets the property name for the specified control state.
        /// For example, to generate code for a SetStateAction on a TreeItem like: -
        ///         myTreeItem.Expanded = true;
        /// this method should return "Expanded" as property name and true as the stateValue.
        /// Otherwise, if this returns null, then the generated code will look like -
        ///         myEdit.SetProperty("State", ControlStates.Expanded);
        /// </summary>
        /// <param name="uiTestControl">The control for which the state property name is required.</param>
        /// <param name="uiState">The state for which the property names are required.</param>
        /// <param name="stateValues">The values for the properties returned.</param>
        /// <returns>The writable property names for the state or null if no property exists.</returns>
        public override string[] GetPropertyForControlState(UITestControl uiTestControl, ControlStates uiState, out bool[] stateValues)
        {
            stateValues = null;
            return null;
        }

        /// <summary>
        /// Gets the class type of where the constants for the property names are defined.
        /// </summary>
        /// <param name="uiTestControl">The control for which the property defining type is required.</param>
        /// <returns>The class which defines the constants for the property names.</returns>
        public override Type GetPropertyNamesClassType(UITestControl uiTestControl)
        {
            return null;
        }

        #endregion

        #region Private Members

        /// <summary>
        /// Initializes the properties dictionary.
        /// </summary>
        /// <returns>The dictionary.</returns>
        private static Dictionary<string, UITestPropertyDescriptor> InitializePropertiesMap()
        {
            Dictionary<string, UITestPropertyDescriptor> map = new Dictionary<string, UITestPropertyDescriptor>(StringComparer.OrdinalIgnoreCase);
            UITestPropertyAttributes read = UITestPropertyAttributes.Readable | UITestPropertyAttributes.DoNotGenerateProperties;
            UITestPropertyAttributes readWrite = read | UITestPropertyAttributes.Writable;
            UITestPropertyAttributes readSearch = read | UITestPropertyAttributes.Searchable;

            map.Add(PropertyNames.WorksheetName, new UITestPropertyDescriptor(typeof(string), read));
            map.Add(PropertyNames.RowIndex, new UITestPropertyDescriptor(typeof(int), readSearch));
            map.Add(PropertyNames.ColumnIndex, new UITestPropertyDescriptor(typeof(int), readSearch));
            map.Add(PropertyNames.Value, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.Text, new UITestPropertyDescriptor(typeof(string), read));
            map.Add(PropertyNames.WidthInChars, new UITestPropertyDescriptor(typeof(double), readWrite));
            map.Add(PropertyNames.HeightInPoints, new UITestPropertyDescriptor(typeof(double), readWrite));
            map.Add(PropertyNames.Formula, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.WrapText, new UITestPropertyDescriptor(typeof(bool), readWrite));

            return map;
        }

        /// <summary>
        /// Gets the cell info from the actual UITestControl.
        /// </summary>
        /// <param name="uiTestControl">The UITestControl.</param>
        /// <returns>The cell info.</returns>
        private static ExcelCellInfo GetCellInfo(UITestControl uiTestControl)
        {
            // Get the ExcelCellElement first.
            ExcelCellElement cellElement = uiTestControl.GetProperty(UITestControl.PropertyNames.UITechnologyElement) as ExcelCellElement;
            if (cellElement != null)
            {
                return cellElement.CellInfo;
            }

            return null;
        }

        /// <summary>
        /// Dictionary of the properties.
        /// </summary>
        private static Dictionary<string, UITestPropertyDescriptor> propertiesMap = InitializePropertiesMap();

        #endregion
    }
}
