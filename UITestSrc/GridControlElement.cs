using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Syncfusion.UITest.GridCommunication;
using System.Windows.Forms;
using Accessibility;
using System.Diagnostics;

namespace Syncfusion.Grid.WPF.UITest
{
    [ComVisible(true)]
    public class GridControlCoreElement : UITechnologyElement
    {
        #region Unused methods
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
            throw new NotSupportedException();
        }

        /// <summary>
        /// Gets a clickable point for this element.  The framework will use
        /// this to get clickable point if UITechnologyElement.GetOption(UITechnologyElementOption.GetClickablePointFrom)
        /// returns GetClickablePointFromTechnologyManager. To use the default algorithm
        /// provided by the framework, throw NotSupportedException.
        /// </summary>
        /// <param name="pointX">The x-coordinate of clickable point.</param>
        /// <param name="pointY">The y-coordinate of clickable point.</param>
        /// <exception cref="System.NotSupportedException">Throws System.NotSupportedException
        /// if this operation is not supported.</exception>
        /// <seealso cref="UITechnologyElementOption.GetClickablePointFrom"/>
        public override void GetClickablePoint(out int pointX, out int pointY)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the QueryId for the related element specified by UITestElementKind.
        /// </summary>
        /// <param name="relatedElement">The kind of related element.</param>
        /// <param name="additionalInfo">Any additional information required.
        /// For example, when relatedElement is UITestElementKind.Child, this gives the name of the child.</param>
        /// <param name="maxDepth">The maximum depth to search under this element to find the required component.</param>
        /// <returns>The QueryId of the element.</returns>
        public override string GetQueryIdForRelatedElement(UITestElementKind relatedElement, object additionalInfo, out int maxDepth)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the amount scrolled in percentage.
        /// </summary>
        /// <param name="srollDirection">The direction for which data is required.</param>
        /// <param name="scrollElement">The element which is either the vertical or horizontal scroll bar.</param>
        /// <returns>The amount in percentage.</returns>
        /// <seealso cref="InitializeProgrammaticScroll"/>
        public override int GetScrolledPercentage(ScrollDirection scrollDirection, IUITechnologyElement scrollElement)
        {
            // Should never get called because InitializeProgrammaticScroll() returns false.
            throw new NotImplementedException();
        }

        /// <summary>
        /// Performs programmatic action, based on the ProgrammaticActionOption passed, on this element.
        /// </summary>
        /// <param name="programmaticActionOption">The option corresponding the action to perform.</param>
        public override void InvokeProgrammaticAction(ProgrammaticActionOption programmaticActionOption)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Does the programmatic scrolling for this element.
        /// </summary>
        /// <param name="scrollDirection">The direction to scroll.</param>
        /// <param name="scrollAmount">The amount to scroll.</param>
        /// <seealso cref="InitializeProgrammaticScroll"/>
        public override void ScrollProgrammatically(ScrollDirection scrollDirection, ScrollAmount scrollAmount)
        {
            // Should never get called because InitializeProgrammaticScroll() returns false.
            throw new NotImplementedException();
        }

        #endregion

        /// <summary>
        /// Caches all the common properties of this element for future use so that these
        /// properties can be used later even when the underlining UI control no longer exists.
        /// This typically includes properties like Name, ClassName, ControlType, QueryId
        /// and other properties used in identification string.
        /// </summary>
        public override void CacheProperties()
        {
        }

        /// <summary>
        /// Gets the 0-based position in the parent element's collection.
        /// </summary>
        public override int ChildIndex
        {
            get { return 0; }
        }

        /// <summary>
        /// Gets the class name of this element.
        /// </summary>
        public override string ClassName
        {
            get
            {
                return Utilities.GridControlClassName;
            }
        }

        /// <summary>
        /// Gets the universal control type of this element.
        /// </summary>
        public override string ControlTypeName
        {
            get
            {
                return ControlType.Window.Name;
            }
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="element">The object to compare with the current object.</param>
        /// <returns>True if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(IUITechnologyElement element)
        {
            var otherElement = element as GridControlCoreElement;
            if (otherElement != null)
            {
                return string.Equals(this.Name, otherElement.Name, StringComparison.Ordinal);
            }

            return false;
        }

        /// <summary>
        /// Gets the user-friendly name for this element like display text that
        /// will help the user to quickly recognize the element on the screen. 
        /// </summary>
        public override string FriendlyName
        {
            get { return this.Name; }
        }

        /// <summary>
        /// Gets the coordinates of the rectangle that completely encloses this element.
        /// </summary>
        /// <remarks>This is in screen coordinates and never cached.</remarks>
        public override void GetBoundingRectangle(out int left, out int top, out int width, out int height)
        {
            left = top = width = height = -1;

            Utilities.RECT rect;
            if (Utilities.GetWindowRect(this.WindowHandle, out rect))
            {
                left = rect.left;
                top = rect.top;
                width = rect.right - rect.left;
                height = rect.bottom - rect.top;
            }
        }

        /// <summary>
        /// Gets the native control type of this element. This can be used in
        /// tandem with the universal type got via GetControlType() in cases
        /// where just the ControlType is not enough to differentiate a control.
        /// For example, if the native technology element is HTML, this could be the tag name.
        /// </summary>
        /// <param name="nativeControlTypeKind">The kind of the native control type requested.</param>
        /// <returns>If supported, the native type of the control or else null.</returns>
        public override object GetNativeControlType(NativeControlTypeKind nativeControlTypeKind)
        {
            if (nativeControlTypeKind == NativeControlTypeKind.AsString)
            {
                return this.ControlTypeName;
            }

            return null;
        }

        /// <summary>
        /// Gets the value for the specified property for this element.
        /// </summary>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns>The value of the property.</returns>
        public override object GetPropertyValue(string propertyName)
        {
            // At the very least, all the properties used in QueryId should be supported here.
            if (string.Equals(PropertyNames.ControlType, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.ControlTypeName;
            }
            else if (string.Equals(PropertyNames.ClassName, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.ClassName;
            }
            else if (string.Equals(PropertyNames.Name, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.Name;
            }
            var pInfos = this.GetType().GetProperties();
            var pInfo = pInfos.FirstOrDefault(p => p.Name == propertyName);
            if (pInfo != null)
            {
                return pInfo.GetValue(this, null);
            }

            return null;
        }

        /// <summary>
        /// Gets the current state information of this element for the given requested states.
        /// If the element does not support querying only the selective states, it can
        /// return the complete state information.
        /// </summary>
        /// <param name="requestedState">The states for which to check.</param>
        /// <returns>The information about the given requested state or complete state information.</returns>
        public override AccessibleStates GetRequestedState(AccessibleStates requestedState)
        {
            IAccessible accessible = Utilities.AccessibleObjectFromWindow(this.WindowHandle);
            if (accessible != null)
            {
                object state = accessible.accState;
                if (state is AccessibleStates)
                {
                    return (AccessibleStates)state;
                }
            }

            return AccessibleStates.Default;
        }

        /// <summary>
        /// Gets the true/false value for right to left format based on the kind specified.
        /// </summary>
        /// <param name="rightToLeftKind">Either the layout or text kind to check for.</param>
        /// <returns>True if layout or text based on the RightToLeftKind passed is right to left,
        /// false otherwise.</returns>
        public override bool GetRightToLeftProperty(RightToLeftKind rightToLeftKind)
        {
            return false;
        }

        /// <summary>
        /// Initializes this element to do programmatic scrolling.
        /// </summary>
        /// <returns>True if element supports programmatic scrolling and
        /// initialization is successful, false otherwise.</returns>
        public override bool InitializeProgrammaticScroll()
        {
            return false;
        }

        /// <summary>
        /// Gets whether this element is a leaf node (i.e. does not have any children) or not.
        /// </summary>
        public override bool IsLeafNode
        {
            get { return false; }
        }

        /// <summary>
        /// Gets a value that indicates whether this element contains protected content or not.
        /// </summary>
        public override bool IsPassword
        {
            get { return false; }
        }

        /// <summary>
        /// Gets whether the tree switching is required for window-less tree switching cases.
        /// </summary>
        /// <remarks>
        /// An example of this would be an ActiveX control hosted inside the browser.
        /// The technology manager of the browser should return true when queried about this
        /// property for HTML OBJECT tag. This will allow the framework to switch to a different
        /// technology manager to support the hosted ActiveX control.
        /// </remarks>
        public override bool IsTreeSwitchingRequired
        {
            get { return false; }
        }

        /// <summary>
        /// Gets the name of this element.
        /// </summary>
        public override string Name
        {
            get
            {
                return Utilities.GetWindowText(this.WindowHandle);
            }
        }

        /// <summary>
        /// Gets the underlying native technology element (like IAccessible) corresponding this element.
        /// </summary>
        public override object NativeElement
        {
            get
            {
                return this.WindowHandle;
            }
        }

        protected IQueryElement queryId;

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
            get
            {
                if (this.queryId == null)
                {
                    this.queryId = new QueryElement();
                    this.queryId.Condition = new AndCondition(
                            new PropertyCondition(PropertyNames.ControlType, this.ControlTypeName),
                            new PropertyCondition(PropertyNames.Name, this.Name));
                }

                return this.queryId;
            }
        }

        /// <summary>
        /// Sets the focus on this element.
        /// </summary>
        public override void SetFocus()
        {
            // no op
        }

        /// <summary>
        /// Gets or sets the container element if one technology is hosted inside another technology.
        /// This is used by the framework.
        /// </summary>
        public override IUITechnologyElement SwitchingElement
        {
            get;
            set;
        }

        protected GridTechnologyManager technologyManager;
        public override UITechnologyManager TechnologyManager
        {
            get
            {
                return this.technologyManager;
            }
        }

        /// <summary>
        /// Gets the name of the corresponding technology.
        /// This value should be same as UITechnologyManager.TechnologyName.
        /// </summary>
        public override string TechnologyName
        {
            get
            {
                return Utilities.GridControlTechnologyName;
            }
        }

        /// <summary>
        /// Gets or sets the top level window corresponding to this element. The top level windows
        /// are typically children of desktop. If this is not set, the framework will set this to
        /// the top-most ancestor of the element (after ignoring the desktop as ancestor).
        /// </summary>
        /// <returns>The top level window.</returns>
        public override UITechnologyElement TopLevelElement
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the value of this element.
        /// </summary>
        public override string Value
        {
            get;
            set;
        }

        /// <summary>
        /// Waits for the element to be ready for user action.
        /// </summary>
        /// <seealso cref="UITechnologyElementOption.WaitForReadyOptions"/>
        /// <exception cref="System.TimeoutException">
        /// Throws TimeoutException if control is not ready in alloted time.
        /// </exception>
        public override void WaitForReady()
        {
            // no op
        }

        protected IntPtr windowHandle;
        /// <summary>
        /// Gets the handle to the Win32 window containing this element.
        /// </summary>
        public override IntPtr WindowHandle
        {
            get
            {
                return this.windowHandle;
            }
        }

        /// <summary>
        /// Gets the children of this control in this technology matching the given condition. 
        /// </summary>
        /// <param name="condition">The condition to match.</param>
        /// <returns>The enumerator for children.</returns>
        internal virtual System.Collections.IEnumerator GetChildren(AndCondition condition)
        {
            int row = 0, col = 0;
            var rowVal = condition.GetPropertyValue(PropertyNames.RowIndex) as string;
            var colVal = condition.GetPropertyValue(PropertyNames.ColumnIndex) as string;
            var name = condition.GetPropertyValue(PropertyNames.GridName) as string;
            if (int.TryParse(rowVal, out row) && int.TryParse(colVal, out col))
            {
#if DEBUG
                Debug.WriteLine(string.Format("Row {0} / Col {1}", row, col));
#endif
                var cellElement = this.technologyManager.GetGridElement(this.windowHandle, new GridCellInfo(row, col, name));
                return new UITechnologyElement[] { cellElement }.GetEnumerator();
            }

            return null;
        }

        protected UITechnologyElement parent = null;

        internal virtual UITechnologyElement Parent
        {
            get
            {
                return this.parent;
            }
        }

        internal GridControlCoreElement(IntPtr windowHandle, GridTechnologyManager manager)
        {
            this.windowHandle = windowHandle;
            this.technologyManager = manager;
        }

    }

    [ComVisible(true)]
    public class GridControlElement : UITechnologyElementProxy
    {

        /// <summary>
        /// Gets the 0-based position in the parent element's collection.
        /// </summary>
        public override int ChildIndex
        {
            get { return 0; }
        }

        /// <summary>
        /// Gets the class name of this element.
        /// </summary>
        public override string ClassName
        {
            get
            {
                return Utilities.GridControlClassName;
            }
        }

        /// <summary>
        /// Gets the universal control type of this element.
        /// </summary>
        public override string ControlTypeName
        {
            get
            {
                return ControlType.Window.Name;
            }
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="element">The object to compare with the current object.</param>
        /// <returns>True if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(IUITechnologyElement element)
        {
            var otherElement = element as GridControlElement;
            if (otherElement != null)
            {
                return string.Equals(this.Name, otherElement.Name, StringComparison.Ordinal);
            }

            return false;
        }

        /// <summary>
        /// Gets the user-friendly name for this element like display text that
        /// will help the user to quickly recognize the element on the screen. 
        /// </summary>
        public override string FriendlyName
        {
            get { return this.Name; }
        }

        /// <summary>
        /// Gets the coordinates of the rectangle that completely encloses this element.
        /// </summary>
        /// <remarks>This is in screen coordinates and never cached.</remarks>
        public override void GetBoundingRectangle(out int left, out int top, out int width, out int height)
        {
            left = top = width = height = -1;

            Utilities.RECT rect;
            if (Utilities.GetWindowRect(this.WindowHandle, out rect))
            {
                left = rect.left;
                top = rect.top;
                width = rect.right - rect.left;
                height = rect.bottom - rect.top;
            }
        }

        /// <summary>
        /// Gets the native control type of this element. This can be used in
        /// tandem with the universal type got via GetControlType() in cases
        /// where just the ControlType is not enough to differentiate a control.
        /// For example, if the native technology element is HTML, this could be the tag name.
        /// </summary>
        /// <param name="nativeControlTypeKind">The kind of the native control type requested.</param>
        /// <returns>If supported, the native type of the control or else null.</returns>
        public override object GetNativeControlType(NativeControlTypeKind nativeControlTypeKind)
        {
            if (nativeControlTypeKind == NativeControlTypeKind.AsString)
            {
                return this.ControlTypeName;
            }

            return null;
        }

        /// <summary>
        /// Gets the value for the specified property for this element.
        /// </summary>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns>The value of the property.</returns>
        public override object GetPropertyValue(string propertyName)
        {
            // At the very least, all the properties used in QueryId should be supported here.
            if (string.Equals(PropertyNames.ControlType, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.ControlTypeName;
            }
            else if (string.Equals(PropertyNames.ClassName, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.ClassName;
            }
            else if (string.Equals(PropertyNames.Name, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return this.Name;
            }
            var pInfos = this.GetType().GetProperties();
            var pInfo = pInfos.FirstOrDefault(p => p.Name == propertyName);
            if (pInfo != null)
            {
                return pInfo.GetValue(this, null);
            }

            return null;
        }

        /// <summary>
        /// Gets the current state information of this element for the given requested states.
        /// If the element does not support querying only the selective states, it can
        /// return the complete state information.
        /// </summary>
        /// <param name="requestedState">The states for which to check.</param>
        /// <returns>The information about the given requested state or complete state information.</returns>
        public override AccessibleStates GetRequestedState(AccessibleStates requestedState)
        {
            IAccessible accessible = Utilities.AccessibleObjectFromWindow(this.WindowHandle);
            if (accessible != null)
            {
                object state = accessible.accState;
                if (state is AccessibleStates)
                {
                    return (AccessibleStates)state;
                }
            }

            return AccessibleStates.Default;
        }

        /// <summary>
        /// Gets the true/false value for right to left format based on the kind specified.
        /// </summary>
        /// <param name="rightToLeftKind">Either the layout or text kind to check for.</param>
        /// <returns>True if layout or text based on the RightToLeftKind passed is right to left,
        /// false otherwise.</returns>
        public override bool GetRightToLeftProperty(RightToLeftKind rightToLeftKind)
        {
            return false;
        }

        /// <summary>
        /// Initializes this element to do programmatic scrolling.
        /// </summary>
        /// <returns>True if element supports programmatic scrolling and
        /// initialization is successful, false otherwise.</returns>
        public override bool InitializeProgrammaticScroll()
        {
            return false;
        }

        /// <summary>
        /// Gets whether this element is a leaf node (i.e. does not have any children) or not.
        /// </summary>
        public override bool IsLeafNode
        {
            get { return false; }
        }

        /// <summary>
        /// Gets a value that indicates whether this element contains protected content or not.
        /// </summary>
        public override bool IsPassword
        {
            get { return false; }
        }

        /// <summary>
        /// Gets whether the tree switching is required for window-less tree switching cases.
        /// </summary>
        /// <remarks>
        /// An example of this would be an ActiveX control hosted inside the browser.
        /// The technology manager of the browser should return true when queried about this
        /// property for HTML OBJECT tag. This will allow the framework to switch to a different
        /// technology manager to support the hosted ActiveX control.
        /// </remarks>
        public override bool IsTreeSwitchingRequired
        {
            get { return false; }
        }

        /// <summary>
        /// Gets the name of this element.
        /// </summary>
        public override string Name
        {
            get
            {
                return Utilities.GetWindowText(this.WindowHandle);
            }
        }

        /// <summary>
        /// Gets the underlying native technology element (like IAccessible) corresponding this element.
        /// </summary>
        public override object NativeElement
        {
            get
            {
                return this.WindowHandle;
            }
        }

        protected IQueryElement queryId;

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
            get
            {
                if (this.queryId == null)
                {
                    this.queryId = new QueryElement();
                    this.queryId.Condition = new AndCondition(
                            new PropertyCondition(PropertyNames.ControlType, this.ControlTypeName),
                            new PropertyCondition(PropertyNames.Name, this.Name));
                }

                return this.queryId;
            }
        }

        protected GridTechnologyManager technologyManager;

        /// <summary>
        /// Gets or sets the top level window corresponding to this element. The top level windows
        /// are typically children of desktop. If this is not set, the framework will set this to
        /// the top-most ancestor of the element (after ignoring the desktop as ancestor).
        /// </summary>
        /// <returns>The top level window.</returns>
        public override UITechnologyElement TopLevelElement
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the value of this element.
        /// </summary>
        public override string Value
        {
            get;
            set;
        }


        protected IntPtr windowHandle;
        /// <summary>
        /// Gets the handle to the Win32 window containing this element.
        /// </summary>
        public override IntPtr WindowHandle
        {
            get
            {
                return this.windowHandle;
            }
        }

        /// <summary>
        /// Gets the children of this control in this technology matching the given condition. 
        /// </summary>
        /// <param name="condition">The condition to match.</param>
        /// <returns>The enumerator for children.</returns>
        internal virtual System.Collections.IEnumerator GetChildren(AndCondition condition)
        {
            int row = 0, col = 0;
            var rowVal = condition.GetPropertyValue(PropertyNames.RowIndex) as string;
            var colVal = condition.GetPropertyValue(PropertyNames.ColumnIndex) as string;
            var GridName = condition.GetPropertyValue(PropertyNames.GridName) as string;
            if (int.TryParse(rowVal, out row) && int.TryParse(colVal, out col))
            {
#if DEBUG
                Debug.WriteLine(string.Format("Row {0} / Col {1}", row, col));
#endif
                var cellElement = this.technologyManager.GetGridElement(this.windowHandle, new GridCellInfo(row, col, GridName));
                return new UITechnologyElement[] { cellElement }.GetEnumerator();
            }

            return null;
        }

        protected UITechnologyElement parent = null;

        internal virtual UITechnologyElement Parent
        {
            get
            {
                return this.parent;
            }
        }

        internal GridControlElement(IntPtr windowHandle, GridTechnologyManager manager)
            : base(manager, new GridControlCoreElement(windowHandle, manager))
        {
            this.windowHandle = windowHandle;
            this.technologyManager = manager;
        }
    }

}
