using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UITest.Common;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;

namespace Syncfusion.Grid.WPF.UITest
{
    internal class CommonPropertyProvider : UITestPropertyProvider
    {
        /// <summary>
        /// Gets the support level of the provider for the specified control.
        /// </summary>
        /// <param name="uiTestControl">The control for which the support level is required.</param>
        /// <returns>The support level for the control.</returns>
        public override int GetControlSupportLevel(UITestControl uiTestControl)
        {
            if (!isInThisProvider &&
                string.Equals(uiTestControl.TechnologyName, Utilities.GridControlTechnologyName, StringComparison.OrdinalIgnoreCase))
            {
                // If technology name is GridControl, we handle it!
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
            isInThisProvider = true;
            try
            {
                UITestControl wpfControl = Utilities.GetCopiedUiaControl(uiTestControl);
                var names = GetInnerProvider(wpfControl).GetPropertyNames(wpfControl);
                return names;
            }
            finally
            {
                isInThisProvider = false;
            }
        }

        /// <summary>
        /// Gets the property descriptor for the given property of the control.
        /// </summary>
        /// <param name="uiTestControl">The control for which descriptor is required.</param>
        /// <param name="propertyName">The name of the property for which descriptor is required.</param>
        /// <returns>The property descriptor of the property.</returns>
        public override UITestPropertyDescriptor GetPropertyDescriptor(UITestControl uiTestControl, string propertyName)
        {
            isInThisProvider = true;
            try
            {
                UITestControl wpfControl = Utilities.GetCopiedUiaControl(uiTestControl);
                var descriptor = GetInnerProvider(wpfControl).GetPropertyDescriptor(wpfControl, propertyName);
                return descriptor;
            }
            finally
            {
                isInThisProvider = false;
            }
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
        public override object GetPropertyValue(UITestControl uiTestControl, string propertyName)
        {
            if (IsUITestControlProperty(propertyName)) throw new NotSupportedException();

            isInThisProvider = true;
            try
            {
                UITestControl wpfControl = Utilities.GetLiveUiaControl(uiTestControl);
                return wpfControl.GetProperty(propertyName);
            }
            finally
            {
                isInThisProvider = false;
            }
        }

        /// <summary>
        /// Sets the property value of the control.
        /// This method is called for all the settable properties defined in UITestControl.PropertyNames
        /// to let the provider override the values. The provider should either support setting of those
        /// properties properly or throw System.NotSupportedException.
        /// </summary>
        /// <param name="uiTestControl">The control for which to set the value.</param>
        /// <param name="propertyName">The name of the property to set.</param>
        /// <param name="propertyValue">The value of the property.</param>
        public override void SetPropertyValue(UITestControl uiTestControl, string propertyName, object propertyValue)
        {
            if (IsUITestControlProperty(propertyName)) throw new NotSupportedException();

            isInThisProvider = true;
            try
            {
                UITestControl wpfControl = Utilities.GetLiveUiaControl(uiTestControl);
                wpfControl.SetProperty(propertyName, propertyValue);
            }
            finally
            {
                isInThisProvider = false;
            }
        }

        #region Code Generation Methods - Not Supported.

        /// <summary>
        /// Gets the specialized class to use for the specified control.
        /// </summary>
        /// <param name="uiTestControl">The control for which a specialized class is to needed.</param>
        /// <returns>Type object of the specialized class.</returns>
        public override Type GetSpecializedClass(UITestControl uiTestControl)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Gets the class type of where the constants for the property names are defined.
        /// </summary>
        /// <param name="uiTestControl">The control for which the property defining type is required.</param>
        /// <returns>The class which defines the constants for the property names.</returns>
        public override Type GetPropertyNamesClassType(UITestControl uiTestControl)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Gets the search properties present by default in the specified specialized class.
        /// </summary>
        /// <param name="specializedClassType">System.Type object of the specialized class.</param>
        /// <returns>Properties already present in the specialized class.</returns>
        public override string[] GetPredefinedSearchProperties(Type specializedClassType)
        {
            throw new NotSupportedException();
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
        public override string GetPropertyForAction(UITestControl uiTestControl, UITestAction action)
        {
            throw new NotSupportedException();
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
            throw new NotSupportedException();
        }

        #endregion

        #region Private Members

        /// <summary>
        /// Gets the UIA provider as the inner provider.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns>The UIA provider.</returns>
        private static UITestPropertyProvider GetInnerProvider(UITestControl control)
        {
            if (!Playback.IsInitialized)
            {
                throw new InvalidOperationException();
            }

            return Playback.GetCorePropertyProvider(control);
        }

        /// <summary>
        /// Is this a UITestControl property?
        /// </summary>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns>True if this is UITestControl property, false otherwise.</returns>
        private static bool IsUITestControlProperty(string propertyName)
        {
            if (uiTestControlProperties == null)
            {
                uiTestControlProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                uiTestControlProperties.Add(UITestControl.PropertyNames.BoundingRectangle, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.ClassName, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.ControlType, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.Enabled, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.Exists, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.FriendlyName, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.HasFocus, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.Instance, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.IsTopParent, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.MaxDepth, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.Name, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.NativeElement, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.QueryId, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.TechnologyName, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.TopParent, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.UITechnologyElement, null);
                uiTestControlProperties.Add(UITestControl.PropertyNames.WindowHandle, null);
            }

            return uiTestControlProperties.ContainsKey(propertyName);
        }

        /// <summary>
        /// Whether the control is already in this provider.
        /// This is to avoid recursion - ideally this should be accounted by the design
        /// and user should not have to do this.
        /// </summary>
        private bool isInThisProvider;

        /// <summary>
        /// Dictionary for UITestControl properties.
        /// </summary>
        private static Dictionary<string, string> uiTestControlProperties;

        #endregion
    }
}

