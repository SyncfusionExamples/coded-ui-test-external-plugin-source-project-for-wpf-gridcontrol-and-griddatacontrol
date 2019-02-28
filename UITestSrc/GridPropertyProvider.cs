using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UITesting;
using Syncfusion.UITest.GridCommunication;
using Microsoft.VisualStudio.TestTools.UITest.Extension;

namespace Syncfusion.Grid.WPF.UITest
{
    internal class GridPropertyProvider : UITestPropertyProvider
    {
        #region Override methods

        /// <summary>
        /// Gets the control support level.
        /// </summary>
        /// <param name="uiTestControl">The UI test control.</param>
        /// <returns></returns>
        public override int GetControlSupportLevel(UITestControl uiTestControl)
        {
            if (string.Equals(uiTestControl.TechnologyName, Utilities.GridControlTechnologyName, StringComparison.OrdinalIgnoreCase) && uiTestControl.ControlType == ControlType.Cell)
            {
                return (int)ControlSupport.ControlSpecificSupport;
            }

            return (int)ControlSupport.NoSupport;
        }

        /// <summary>
        ///  Gets the search properties present by default in the specified specialized class.
        /// </summary>
        /// <param name="specializedClass">System.Type object of the specialized class.</param>
        /// <returns>Properties already present in the specialized class.</returns>
        public override string[] GetPredefinedSearchProperties(Type specializedClass)
        {
            return null;
        }

        /// <summary>
        /// Gets the property descriptor.
        /// </summary>
        /// <param name="uiTestControl">The UI test control.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <returns></returns>
        public override UITestPropertyDescriptor GetPropertyDescriptor(UITestControl uiTestControl, string propertyName)
        {
            if (propertiesMap.ContainsKey(propertyName))
            {
                return propertiesMap[propertyName];
            }

            return null;
        }

        /// <summary>
        /// Gets the property for action.
        /// </summary>
        /// <param name="uiTestControl">The UI test control.</param>
        /// <param name="action">The action.</param>
        /// <returns></returns>
        public override string GetPropertyForAction(UITestControl uiTestControl, Microsoft.VisualStudio.TestTools.UITest.Common.UITestAction action)
        {
            return null;
        }

        /// <summary>
        /// Gets the state of the property for control.
        /// </summary>
        /// <param name="uiTestControl">The UI test control.</param>
        /// <param name="uiState">State of the UI.</param>
        /// <param name="stateValues">The state values.</param>
        /// <returns></returns>
        public override string[] GetPropertyForControlState(UITestControl uiTestControl, Microsoft.VisualStudio.TestTools.UITest.Extension.ControlStates uiState, out bool[] stateValues)
        {
            stateValues = null;
            return null;
        }

        /// <summary>
        /// Gets the PropertyNames for the corresponding uiTestControl
        /// </summary>
        /// <param name="uiTestControl">UITestControl</param>
        /// <returns>PropertyNames</returns>
        public override ICollection<string> GetPropertyNames(UITestControl uiTestControl)
        {
            return propertiesMap.Keys;
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

        public override object GetPropertyValue(UITestControl uiTestControl, string propertyName)
        {
            var technologyElement = uiTestControl.GetProperty("UITechnologyElement") as GridCellElement;
            object result = null;
            if (technologyElement != null)
            {
                if (propertiesMap.ContainsKey(propertyName))
                {
                    var cellInfo = this.GetCellInfo(uiTestControl);
                    result = GridCommunicator.Instance.GetStyleInfoProperty(cellInfo, propertyName);
                }
                else
                {
                    if (propertyName == UITestControl.PropertyNames.BoundingRectangle)
                    {
                        try
                        {
                            int num, num1, num2, num3;
                            technologyElement.GetBoundingRectangle(out num, out num1, out num2, out num3);
                            result = new System.Drawing.Rectangle(num, num1, num2, num3);

                        }
                        catch
                        {
                            result = System.Drawing.Rectangle.Empty;
                        }
                    }
                    else if (propertyName == UITestControl.PropertyNames.FriendlyName)
                    {
                        result = technologyElement.FriendlyName;
                    }
                    else if (propertyName == UITestControl.PropertyNames.NativeElement)
                    {
                        result = technologyElement.NativeElement;
                    }
                    else if (propertyName == UITestControl.PropertyNames.WindowHandle)
                    {
                        result = technologyElement.WindowHandle;
                    }
                    else
                    {
                        result = technologyElement.GetPropertyValue(propertyName);
                    }
                }
            }

            return result;
        }

        public override Type GetSpecializedClass(UITestControl uiTestControl)
        {
            return null;
        }

        public override void SetPropertyValue(UITestControl uiTestControl, string propertyName, object value)
        {
            // Just to match real user behavior closely, set Value & Formula
            // properties by actual typing of the text in the cell.
            if (string.Equals(propertyName, PropertyNames.CellValue, StringComparison.OrdinalIgnoreCase))
            {
                Keyboard.SendKeys(uiTestControl, value as string);
            }

            var cellInfo = this.GetCellInfo(uiTestControl);
            GridCommunicator.Instance.SetStyleInfoProperty(cellInfo, propertyName, value);
        }

        #endregion

        #region Private methods

        private static Dictionary<string, UITestPropertyDescriptor> propertiesMap = InitializePropertiesMap();
        private static Dictionary<string, UITestPropertyDescriptor> InitializePropertiesMap()
        {
            var map = new Dictionary<string, UITestPropertyDescriptor>(StringComparer.OrdinalIgnoreCase);
            UITestPropertyAttributes read = UITestPropertyAttributes.Readable | UITestPropertyAttributes.DoNotGenerateProperties;
            UITestPropertyAttributes readWrite = read | UITestPropertyAttributes.Writable;
            UITestPropertyAttributes readSearch = read | UITestPropertyAttributes.Searchable;

            map.Add(PropertyNames.CellValue, new UITestPropertyDescriptor(typeof(object), readWrite));
            map.Add(PropertyNames.ColumnHeader, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.Text, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.Description, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.SelectedRanges, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.Format, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.FormulaTag, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.CellWidth, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.CellHeight, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.FormattedText, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.RowCount, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.ColumnCount, new UITestPropertyDescriptor(typeof(string), readWrite));
            map.Add(PropertyNames.GridName, new UITestPropertyDescriptor(typeof(string), readWrite));
            return map;
        }

        private GridCellInfo GetCellInfo(UITestControl uiTestControl)
        {
            var cellElement = uiTestControl.GetProperty("UITechnologyElement") as GridCellElement;
            if (cellElement != null)
            {
                return cellElement.CellInfo;
            }

            return null;
        }

        #endregion
    }
}
