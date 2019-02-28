namespace Syncfusion.Grid.WPF.UITest
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.TestTools.UITest.Extension;
    using Syncfusion.UITest.GridCommunication;
    using System.Diagnostics;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using System.Windows;
    using System.Collections;

    [ComVisible(true)]
    public sealed class GridTechnologyManager : UITechnologyManagerProxy
    {
        public GridTechnologyManager()
            : base("UIA", Utilities.GridControlTechnologyName)
        {
        }

        public IUITechnologyElement coreelement;

        /// <summary>
        /// Converts the given element of another technology to new element of this technology manager.
        /// This is used for operations like switching between hosted and hosting technologies.
        /// </summary>
        /// <param name="elementToConvert">The element to convert.</param>
        /// <param name="supportLevel">The level of support provided for the
        /// converted element by this technology manager.</param>
        /// <returns>The new converted element in this technology or null if no conversion is possible.</returns>
        /// <seealso cref="GetControlSupportLevel"/>
        /// <seealso cref="ControlSupport"/>
        public override IUITechnologyElement ConvertToThisTechnology(IUITechnologyElement elementToConvert, out int supportLevel)
        {
            supportLevel = (int)ControlSupport.NoSupport;
            coreelement = elementToConvert;
            if (elementToConvert is GridControlElement)
            {
                supportLevel = (int)ControlSupport.ControlSpecificSupport;
                return elementToConvert;
            }
            else
            {
                var windowHandle = elementToConvert.WindowHandle;
                if (Utilities.IsGridControlWindow(windowHandle))
                {
                    supportLevel = (int)ControlSupport.ControlSpecificSupport;
                    return GetGridElement(windowHandle, null);
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the support level of this technology manager for the elements(s) in the given window.
        /// The framework uses this function to select the right technology manager for the element. 
        /// </summary>
        /// <param name="windowHandle">The window handle of the element.</param>
        /// <returns>An integer that indicates the level of support provided for the element
        /// by this technology manager. The higher the value the stronger the support.</returns>
        /// <seealso cref="ControlSupport"/>
        public override int GetControlSupportLevel(IntPtr windowHandle)
        {
            if (!this.isRecordingSession && Utilities.IsGridControlWindow(windowHandle))
            {
#if DEBUG
                Debug.WriteLine("Control Specific Support");
#endif
                return (int)ControlSupport.ControlSpecificSupport;
            }
            else
            {
                return (int)ControlSupport.NoSupport;
            }
        }

        /// <summary>
        /// Gets the element from the given native (underlying) technology element.
        /// </summary>
        /// <param name="nativeElement">The native technology element (like IAccessible).</param>
        /// <returns>The IUITechnologyElement from the native element.</returns>
        /// <seealso cref="UITechnologyElement.NativeElement"/>
        public override IUITechnologyElement GetElementFromNativeElement(object nativeElement)
        {
            object[] parts = nativeElement as object[];
            if (parts != null && parts.Length == 2 && parts[0] is IntPtr && parts[1] is GridInfo)
            {
                // Get the cell or worksheet as appropriate.
                IntPtr windowHandle = (IntPtr)parts[0];
                GridInfo elementInfo = (GridInfo)parts[1];
                return this.GetGridElement(windowHandle, elementInfo);
            }
            else if (nativeElement is IntPtr)
            {
                // For window handle, get the Excel worksheet.
                return GetElementFromWindowHandle((IntPtr)nativeElement);
            }

            return null;
        }

        /// <summary>
        /// Gets the element at the given screen coordinates.
        /// </summary>
        /// <param name="pointX">The x-coordinate of the screen location.</param>
        /// <param name="pointY">The y-coordinate of the screen location.</param>
        /// <returns>The IUITechnologyElement at the screen coordinates specified.</returns>
        public override IUITechnologyElement GetElementFromPoint(int pointX, int pointY)
        {
            // RPC method that returns the element from Point
            var windowHandle = Utilities.WindowFromPoint(pointX, pointY);
            var element = GridCommunicator.Instance.GetElementFromPoint(pointX, pointY);
            if (element == null)
            {
                var element1 = (Playback.GetCoreTechnologyManager("UIA")).GetElementFromPoint(pointX, pointY);
                if (Utilities.IsUiaElement(element1))
                    return new CommonControlElement(windowHandle, this, element1);
                return null;
            }
            return this.GetGridElement(windowHandle, element);
        }

        /// <summary>
        /// Retrieves the element identified by the provided window handle.
        /// </summary>
        /// <param name="handle">A <see cref="T:System.IntPtr"/> that identifies an existing element.</param>
        /// <returns>The specified element.</returns>
        public override IUITechnologyElement GetElementFromWindowHandle(IntPtr handle)
        {
            if (Utilities.IsGridControlWindow(handle))
            {
                return new CommonControlElement(handle, this, (Playback.GetCoreTechnologyManager("UIA").GetElementFromWindowHandle(handle)));
            }

            return null;
        }

        /// <summary>
        /// Gets the focused element i.e. the element that will receive keyboard events at this instance.
        /// </summary>
        /// <param name="handle">The handle of the window which has the focus.
        /// The element which has focus could be the window itself or a
        /// descendant of this window.</param>
        /// <returns>The element that has the focus or null if there is no element with focus.</returns>
        public override IUITechnologyElement GetFocusedElement(IntPtr handle)
        {
            // TODO - Implement a method call in the RPC
            return this.GetGridElement(handle, null);
        }

        /// <summary>
        /// Returns a <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUISynchronizationWaiter"/> using the provided element and event type.
        /// </summary>
        /// <param name="element">An <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUITechnologyElement"/> object.</param>
        /// <param name="eventType">A <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.UITestEventType"/> object for which a waiter is required.</param>
        /// <returns>The requested waiter.</returns>
        /// <exception cref="T:System.NotSupportedException">Throw only if this method is not supported by this technology manager. </exception>
        public override IUISynchronizationWaiter GetSynchronizationWaiter(IUITechnologyElement element, UITestEventType eventType)
        {
            return null;
        }

        /// <summary>
        /// Matches the element against the parsedQueryIdCookie condition.
        /// </summary>
        /// <param name="element">The element to match against the conditions.</param>
        /// <param name="parsedQueryIdCookie">The cookie of previously parsed QueryId.</param>
        /// <param name="useEngine">
        /// This is set to true by the technology manager if it wants to use
        /// the framework for matching the complete or part of query element.</param>
        /// <returns>True if the element matches the condition, false otherwise.</returns>
        /// <remarks>
        /// This is an optional method and if the technology manager does not support
        /// this method it should throw System.NotSupportedException exception. If the Search()
        /// is not supported then the framework uses GetChildren() API to do breadth-first
        /// traversal and for each element uses MatchElement() API to match & find.
        /// 
        /// Note that a technology has to support either this or Search.
        /// </remarks>
        /// <seealso cref="ParseQueryId"/>
        /// <seealso cref="Search"/>
        public override bool MatchElement(IUITechnologyElement element, object parsedQueryIdCookie, out bool useEngine)
        {
            // Get the condition out of the cookie as set by ParseQueryId() API.
            IQueryCondition condition = parsedQueryIdCookie as AndCondition;
            if (condition != null)
            {
                // Use the Match API to do the matching. Note that this API
                // will call into ExcelElement.GetPropertyValue() method to
                // get the value of the properties for matching. 
                useEngine = false;
                return condition.Match(element);
            }
            else
            {
                useEngine = true;
                return false;
            }
        }

        /// <summary>
        /// Parses the query element string and returns the parsedQueryIdCookie to be
        /// used later during Search() or MatchElement() or GetChildren() call for
        /// either searching or matching or getting children that has the same query string.
        /// </summary>
        /// <param name="queryElement">The query element string to parse.</param>
        /// <param name="parsedQueryIdCookie">The cookie of the parsed QueryId to be used later.</param>
        /// <returns>The remaining part of query element string that is not supported
        /// by this technology manager. The framework may or may not support the remaining part.</returns>
        public override string ParseQueryId(string queryElement, out object parsedQueryIdCookie)
        {
            // Use the AndCondition.Parse() API to get the condition object
            // from string. Here, all the properties possible in query id are
            // supported by this technology manager itself.
            IQueryCondition condition = null;
            try
            {
                condition = AndCondition.Parse(queryElement);
            }
            catch (ArgumentException)
            {
            }

            if (condition == null)
            {
                // Implies parse failed. This should not be the case.
                Debug.Fail("ParseQueryId failed");
                parsedQueryIdCookie = null;
                return queryElement;
            }
            else
            {
                // Store the condition as the cookie to be used later.
                parsedQueryIdCookie = condition;
                return string.Empty;
            }
        }

        public override void ProcessMouseEnter(IntPtr handle)
        {
            // no op
        }

        #region Navigation methods

        /// <summary>
        /// Gets the parent of the given element in the user interface hierarchy.
        /// </summary>
        /// <param name="element">The element whose parent is needed.</param>
        /// <returns>The parent element or null if the element passed is the
        /// root element in this technology.</returns>
        public override IUITechnologyElement GetParent(IUITechnologyElement element)
        {
            coreelement = element;
            var cellElement = element as GridCellElement;
            if (cellElement != null)
            {
                return cellElement.Parent;
            }

            return null;
        }

        /// <summary>
        /// Gets the enumerator for children of the given IUITechnologyElement.
        /// </summary>
        /// <param name="element">The IUITechnologyElement whose child enumerator is needed.</param>
        /// <param name="parsedQueryIdCookie">The cookie of previously parsed QueryId to filter matching children.</param>
        /// <returns>The enumerator for children.</returns>
        /// <seealso cref="ParseQueryId"/>
        public override System.Collections.IEnumerator GetChildren(IUITechnologyElement element, object parsedQueryIdCookie)
        {
            if (parsedQueryIdCookie != null)
            {
                // this conversion is while playback the test casees for grid cell we need to get childrens of DataGrid .
                //so we create instance of GridControlElement  instead of CommonControlElement
                var conditionsStack = ((parsedQueryIdCookie as QueryCondition).Conditions as IQueryCondition[]);

                foreach (var item in conditionsStack)
                {
                    if (!(item is PropertyCondition))
                        continue;

                    var controlType = (item as PropertyCondition).Value;
                    var property = (item as PropertyCondition).PropertyName;
                    if (property == "ControlType" && controlType.ToString() == "Cell" && !(element is GridControlElement))
                    {
                        element = new GridControlElement(element.WindowHandle, this);
                        break;
                    }
                }
            }
            var gridElement = element as GridControlElement;
            var condition = parsedQueryIdCookie as AndCondition;
            coreelement = element;
            if (gridElement != null)
            {
                if (condition != null)
                {
                    return gridElement.GetChildren(condition);
                }
                else
                {
                    // TODO - return all children
                }
            }
            else
                return new ChildrenEnumerator(this, (Playback.GetCoreTechnologyManager("UIA")).GetChildren(Utilities.GetUiaElement(element), parsedQueryIdCookie));

            return null;
        }

        /// <summary>
        /// Gets the next sibling of the given element in the user interface hierarchy.
        /// </summary>
        /// <param name="element">The element whose next sibling is needed.</param>
        /// <returns>The next sibling or null if none is present.</returns>
        public override IUITechnologyElement GetNextSibling(IUITechnologyElement element)
        {
            // TODO - Sibling navigation is required to get the arrow keys of Spy control working. 
            return null;
        }

        /// <summary>
        /// Gets the previous sibling of the given element in the user interface hierarchy.
        /// </summary>
        /// <param name="element">The element whose previous sibling is needed.</param>
        /// <returns>The previous sibling or null if none is present.</returns>
        public override IUITechnologyElement GetPreviousSibling(IUITechnologyElement element)
        {
            // TODO - Sibling navigation is required to get the arrow keys of Spy control working. 
            return null;
        }

        #endregion

        #region Search methods - not required
        /// <summary>
        /// Searches for an element matching the given query element. If the underlying
        /// UI Technology has rich APIs to search/navigate the UI hierarchy then
        /// implementing this method could improve the playback performance significantly.
        /// </summary>
        /// <param name="parsedQueryIdCookie">The cookie of previously parsed QueryId.</param>
        /// <param name="parentElement">The parent element under which to search.</param>
        /// <param name="maxDepth">The maximum tree depth to search.</param>
        /// <returns>An array of matched elements or null if none matched.</returns>
        /// <remarks>
        /// This is an optional method and if the technology manager does not support
        /// this method it should throw System.NotSupportedException exception. If this
        /// is not supported then the framework uses GetChildren() API to do breadth-first
        /// traversal and for each element uses MatchElement() API to match & find.
        /// 
        /// Note that a technology has to support either this or MatchElement.
        /// </remarks>
        /// <seealso cref="ParseQueryId"/>
        /// <seealso cref="MatchElement"/>
        /// <seealso cref="UITechnologyManagerProperty.SearchSupported"/>
        public override object[] Search(object parsedQueryIdCookie, IUITechnologyElement parentElement, int maxDepth)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Cancels any wait or search operation being performed by this technology manager
        /// because of call to WaitForReady or Search methods.
        /// </summary>
        /// <remarks>
        /// This call is made on another thread as both the WaitForReady and Search methods are blocking.
        /// </remarks>
        public override void CancelStep()
        {
            // no op
        }

        /// <summary>
        /// Gets the information about the most recent invocation of the technology manager.
        /// </summary>
        /// <returns>Information about the most recent invocation of the technology manager.</returns>
        public override ILastInvocationInfo GetLastInvocationInfo()
        {
            return (Playback.GetCoreTechnologyManager("UIA")).GetLastInvocationInfo();
            //throw new NotImplementedException();
        }

        #endregion

        #region Add / Remove Event handler

        /// <summary>
        /// Adds an event handler to this technology manager.
        /// </summary>
        /// <param name="element">The <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUITechnologyElement"/> and its descendents for which this event should be raised.</param>
        /// <param name="eventType">A member of the <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.UITestEventType"/> enumeration that specifies the type of event.</param>
        /// <param name="eventSink">An <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUITestEventNotify"/> that logs the events.</param>
        /// <returns>
        /// true if the event type is supported and is successfully added; otherwise, false.
        /// </returns>
        public override bool AddEventHandler(IUITechnologyElement element, UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        /// <summary>
        /// Adds a global event sink to this technology manager.
        /// </summary>
        /// <param name="eventType">A member of the <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.UITestEventType"/> enumeration that specifies the type of event.</param>
        /// <param name="eventSink">An <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUITestEventNotify"/> that logs the events.</param>
        /// <returns>
        /// true if the event type is supported and is successfully added; otherwise, false.
        /// </returns>
        public override bool AddGlobalEventHandler(UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        /// <summary>
        /// Removes the specified event from the given element and all its descendents.
        /// </summary>
        /// <param name="element">An <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUITechnologyElement"/> object.</param>
        /// <param name="eventType">A member of the <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.UITestEventType"/> enumeration.</param>
        /// <param name="eventSink">An <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUITestEventNotify"/> object representing the registered event sink.</param>
        /// <returns>
        /// true if the event was successfully removed; otherwise, false.
        /// </returns>
        public override bool RemoveEventHandler(IUITechnologyElement element, UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        /// <summary>
        /// Removes the specified event.
        /// </summary>
        /// <param name="eventType">A member of the <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.UITestEventType"/> enumeration.</param>
        /// <param name="eventSink">An <see cref="T:Microsoft.VisualStudio.TestTools.UITest.Extension.IUITestEventNotify"/> object representing the registered event sink.</param>
        /// <returns>
        /// true if the event was successfully removed; otherwise, false.
        /// </returns>
        public override bool RemoveGlobalEventHandler(UITestEventType eventType, IUITestEventNotify eventSink)
        {
            return false;
        }

        #endregion

        #region Initialize / Clean up methods - not required for now

        private bool isRecordingSession;
        public override void StartSession(bool recordingSession)
        {
            this.isRecordingSession = recordingSession;
        }

        public override void StopSession()
        {
            // no op
        }

        #endregion

        #region Internal / Private method

        internal UITechnologyElement GetGridElement(IntPtr windowHandle, GridInfo elementInfo)
        {
            if (elementInfo is GridCellInfo)
            {
                var cellInfo = elementInfo as GridCellInfo;
                if (cellInfo.Parent == null)
                {
                    cellInfo.Parent = new GridControlInfo();
                }
                return new GridCellElement(windowHandle, elementInfo as GridCellInfo, this);
            }

            return new GridControlElement(windowHandle, this);
        }

        #endregion

        public override IUITechnologyElement ConvertToExtensionElement(IUITechnologyElement coreElement)
        {
            coreelement = coreElement;
#if SyncfusionFramework4_5_1 || SyncfusionFramework4_6
            //From Visual studio update  2 they provided support for window less control (Windows Store app)
            // Due to this we must implement this override method.this method is executed instead of GetEleemntFromPoint() method.
            var element = GridCommunicator.Instance.GetElementFromPoint(Mouse.Location.X, Mouse.Location.Y);
            if (element == null)
            {
                var element1 = (Playback.GetCoreTechnologyManager("UIA")).GetElementFromPoint(Mouse.Location.X, Mouse.Location.Y);
                if (Utilities.IsUiaElement(element1))
                    return new CommonControlElement(coreelement.WindowHandle, this, element1);
                return null;
            }
            return this.GetGridElement(coreelement.WindowHandle, element);
#else
            throw new NotImplementedException();
#endif
        }
    }

}
