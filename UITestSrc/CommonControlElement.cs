using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UITest.Common;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using System.Runtime.InteropServices;
using System.Collections;

namespace Syncfusion.Grid.WPF.UITest
{
    [ComVisible(true)]
    public class CommonControlElement : UITechnologyElement
    {
        protected IntPtr windowHandle;

        /// <summary>
        /// Gets the handle to the Win32 window containing this element.
        /// </summary>
        public override IntPtr WindowHandle
        {
            get
            {
                if (this.windowHandle == null)
                    return InnerElement.WindowHandle;
                return this.windowHandle;
            }
        }

        /// <summary>
        /// Gets the name of the corresponding technology.
        /// This value should be same as UITechnologyManager.TechnologyName.
        /// </summary>
        public override string TechnologyName
        {
            get { return technologyManager.TechnologyName; }
        }

        /// <summary>
        /// Gets the corresponding technology manager.
        /// </summary>
        public override UITechnologyManager TechnologyManager
        {
            get { return technologyManager; }
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="element">The object to compare with the current object.</param>
        /// <returns>True if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(IUITechnologyElement element)
        {
            CommonControlElement other = element as CommonControlElement;
            if ((object)other != null)
            {
                return InnerElement.Equals(other.InnerElement);
            }

            return false;
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="obj">The object to compare with the current object.</param>
        /// <returns>True if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            return Equals(obj as CommonControlElement);
        }

        #region Public Properties - Simply call corresponding property on UIA element

        /// <summary>
        /// Gets the 0-based position in the parent element's collection.
        /// </summary>
        public override int ChildIndex
        {
            get { return InnerElement.ChildIndex; }
        }

        /// <summary>
        /// Gets the class name of this element.
        /// </summary>
        public override string ClassName
        {
            get { return InnerElement.ClassName; }
        }

        /// <summary>
        /// Gets the universal control type of this element.
        /// </summary>
        public override string ControlTypeName
        {
            get { return InnerElement.ControlTypeName; }
        }

        /// <summary>
        /// Gets the user-friendly name for this element like display text that
        /// will help the user to quickly recognize the element on the screen. 
        /// </summary>
        public override string FriendlyName
        {
            get { return InnerElement.FriendlyName; }
        }

        /// <summary>
        /// Gets whether this element is a leaf node (i.e. does not have any children) or not.
        /// </summary>
        public override bool IsLeafNode
        {
            get { return InnerElement.IsLeafNode; }
        }

        /// <summary>
        /// Gets a value that indicates whether this element contains protected content or not.
        /// </summary>
        public override bool IsPassword
        {
            get { return InnerElement.IsPassword; }
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
            get { return InnerElement.IsTreeSwitchingRequired; }
        }

        /// <summary>
        /// Gets the name of this element.
        /// </summary>
        public override string Name
        {
            get { return InnerElement.Name; }
        }

        /// <summary>
        /// Gets the underlying native technology element (like IAccessible) corresponding this element.
        /// </summary>
        public override object NativeElement
        {
            get { return InnerElement.NativeElement; }
        }

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
            get { return InnerElement.QueryId; }
        }

        /// <summary>
        /// Gets or sets the container element if one technology is hosted inside another technology.
        /// This is used by the framework.
        /// </summary>
        public override IUITechnologyElement SwitchingElement
        {
            get
            {
                return InnerElement.SwitchingElement;
            }
            set
            {
                InnerElement.SwitchingElement = value;
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
            get
            {
                return InnerElement.TopLevelElement;
            }
            set
            {
                InnerElement.TopLevelElement = value;
            }
        }

        /// <summary>
        /// Gets the value of this element.
        /// </summary>
        public override string Value
        {
            get
            {
                return InnerElement.Value;
            }
            set
            {
                InnerElement.Value = value;
            }
        }

        #endregion

        #region Public Methods - Simply call corresponding method on UIA element

        /// <summary>
        /// Caches all the common properties of this element for future use so that these
        /// properties can be used later even when the underlining UI control no longer exists.
        /// This typically includes properties like Name, ClassName, ControlType, QueryId
        /// and other properties used in identification string.
        /// </summary>
        public override void CacheProperties()
        {
            InnerElement.CacheProperties();
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
            InnerElement.EnsureVisibleByScrolling(pointX, pointY, ref outpointX, ref outpointY);
        }

        /// <summary>
        /// Gets the coordinates of the rectangle that completely encloses this element.
        /// </summary>
        /// <remarks>This is in screen coordinates and never cached.</remarks>
        public override void GetBoundingRectangle(out int left, out int top, out int width, out int height)
        {
            InnerElement.GetBoundingRectangle(out left, out top, out width, out height);
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
            InnerElement.GetClickablePoint(out pointX, out pointY);
        }

        /// <summary>
        /// Gets the hash code for this object.
        /// .NET Design Guidelines suggests overridding this too if Equals is overridden. 
        /// </summary>
        /// <returns>The hash code.</returns>
        public override int GetHashCode()
        {
            return InnerElement.GetHashCode();
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
            return InnerElement.GetNativeControlType(nativeControlTypeKind);
        }

        /// <summary>
        /// Gets the option for this IUITechnologyElement.
        /// </summary>
        /// <param name="technologyElementOption">The element option to get.</param>
        /// <returns>The value of this element option. </returns>
        /// <exception cref="System.NotSupportedException">Throws System.NotSupportedException
        /// if the element option is not supported.</exception>
        public override object GetOption(UITechnologyElementOption technologyElementOption)
        {
            return InnerElement.GetOption(technologyElementOption);
        }

        /// <summary>
        /// Gets the value for the specified property for this element.
        /// </summary>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns>The value of the property.</returns>
        public override object GetPropertyValue(string propertyName)
        {
            return InnerElement.GetPropertyValue(propertyName);
        }

        /// <summary>
        /// Gets the QueryId for the related element specified by UITestElementKind.
        /// </summary>
        /// <param name="relatedElement">The kind of related element.</param>
        /// <param name="additionalInfo">Any additional information required.
        /// For example, when relatedElement is UITestElementKind.Child, this gives the name of the child.</param>
        /// <param name="maxDepth">The maximum depth to search under this element to find the required component.</param>
        /// <returns>The QueryId of the element.</returns>
        public override string GetQueryIdForRelatedElement(UITestElementKind relatedElement,
            object additionalInfo, out int maxDepth)
        {
            return InnerElement.GetQueryIdForRelatedElement(relatedElement, additionalInfo, out maxDepth);
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
            return InnerElement.GetRequestedState(requestedState);
        }

        /// <summary>
        /// Gets the true/false value for right to left format based on the kind specified.
        /// </summary>
        /// <param name="rightToLeftKind">Either the layout or text kind to check for.</param>
        /// <returns>True if layout or text based on the RightToLeftKind passed is right to left,
        /// false otherwise.</returns>
        public override bool GetRightToLeftProperty(RightToLeftKind rightToLeftKind)
        {
            return InnerElement.GetRightToLeftProperty(rightToLeftKind);
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
            return InnerElement.GetScrolledPercentage(scrollDirection, scrollElement);
        }

        /// <summary>
        /// Initializes this element to do programmatic scrolling.
        /// </summary>
        /// <returns>True if element supports programmatic scrolling and
        /// initialization is successful, false otherwise.</returns>
        public override bool InitializeProgrammaticScroll()
        {
            return InnerElement.InitializeProgrammaticScroll();
        }

        /// <summary>
        /// Performs programmatic action, based on the ProgrammaticActionOption passed, on this element.
        /// </summary>
        /// <param name="programmaticActionOption">The option corresponding the action to perform.</param>
        public override void InvokeProgrammaticAction(ProgrammaticActionOption programmaticActionOption)
        {
            InnerElement.InvokeProgrammaticAction(programmaticActionOption);
        }

        /// <summary>
        /// Does the programmatic scrolling for this element.
        /// </summary>
        /// <param name="scrollDirection">The direction to scroll.</param>
        /// <param name="scrollAmount">The amount to scroll.</param>
        /// <seealso cref="InitializeProgrammaticScroll"/>
        public override void ScrollProgrammatically(ScrollDirection scrollDirection, ScrollAmount scrollAmount)
        {
            InnerElement.ScrollProgrammatically(scrollDirection, scrollAmount);
        }

        /// <summary>
        /// Sets the focus on this element.
        /// </summary>
        public override void SetFocus()
        {
            InnerElement.SetFocus();
        }

        /// <summary>
        /// Sets the option for this IUITechnologyElement.
        /// </summary>
        /// <param name="technologyElementOption">The element option to set.</param>
        /// <param name="optionValue">The value of this element option.</param>
        /// <exception cref="System.NotSupportedException">Throws System.NotSupportedException
        /// if the element option is not supported.</exception>
        public override void SetOption(UITechnologyElementOption technologyElementOption, object optionValue)
        {
            InnerElement.SetOption(technologyElementOption, optionValue);
        }

        /// <summary>
        /// Gets the string representation of this control.
        /// </summary>
        /// <returns>The string representation.</returns>
        public override string ToString()
        {
            return InnerElement.ToString();
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
            InnerElement.WaitForReady();
        }

        #endregion

        /// <summary>
        /// Creates a element from the UIA element.
        /// </summary>
        /// <param name="technologyManager">The technology manager.</param>
        /// <param name="innerElement">The UIA element.</param>
        internal CommonControlElement(IntPtr windowHandle, UITechnologyManager technologyManager, IUITechnologyElement innerElement)
        {
            //if (!Utilities.IsUiaElement(innerElement))
            //{
            //    throw new ArgumentException();
            //}
            this.windowHandle = windowHandle;
            this.innerElement = innerElement as UITechnologyElement;
            this.technologyManager = technologyManager;
        }

        /// <summary>
        /// Gets the inner UIA element.
        /// </summary>
        internal UITechnologyElement InnerElement
        {
            get
            {
                return innerElement;
            }
        }

        private UITechnologyElement innerElement;
        private UITechnologyManager technologyManager;
    }

    /// <summary>
    /// CommonControlElement class to convert elements of UIA technology manager to this technology manager.
    /// </summary>
    internal class ChildrenEnumerator : IEnumerator
    {
        public ChildrenEnumerator(GridTechnologyManager technologyManager, IEnumerator innerEnumerator)
        {
            this.technologyManager = technologyManager;
            this.innerEnumerator = innerEnumerator;
        }

        public object Current
        {
            get
            {
                return new CommonControlElement((this.innerEnumerator.Current as UITechnologyElement).WindowHandle, technologyManager, this.innerEnumerator.Current as UITechnologyElement);
            }
        }

        public bool MoveNext()
        {
            return this.innerEnumerator.MoveNext();
        }

        public void Reset()
        {
            this.innerEnumerator.Reset();
        }

        private GridTechnologyManager technologyManager;
        private IEnumerator innerEnumerator;
    }
}
