using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UITesting;
using Syncfusion.UITest.GridCommunication;

namespace Syncfusion.Grid.WPF.UITest
{
    [ComVisible(true)]
    public class GridCellElement : GridControlElement
    {
        #region Basic properties and methods

        public override int ChildIndex
        {
            get { return this.CellInfo.RowIndex * this.CellInfo.ColumnIndex; }
        }

        public override string ClassName
        {
            get { return "GridControl.Cell"; }
        }

        public override string ControlTypeName
        {
            get { return ControlType.Cell.Name; }
        }

        public override bool IsLeafNode
        {
            get { return true; }
        }

        public override string Name
        {
            get { return this.CellInfo.ToString(); }
        }

        public override object NativeElement
        {
            get { return new object[] { this.WindowHandle, this.CellInfo }; }
        }

        public override void GetBoundingRectangle(out int left, out int top, out int width, out int height)
        {
            left = top = width = height = -1;

            Utilities.RECT windowRect;
            if (Utilities.GetWindowRect(this.WindowHandle, out windowRect))
            {
                var cellRect = GridCommunicator.Instance.GetCellBoundingRect(this.CellInfo);

                #region Comments

                //left = windowRect.left + Utilities.PointToPixel(cellRect.Left, true);
                //top = windowRect.top + Utilities.PointToPixel(cellRect.Top, false);
                //left = Utilities.PointToPixel(cellRect.Left, true);
                //top = Utilities.PointToPixel(cellRect.Top, false);
                //width = Utilities.PointToPixel(cellRect.Width, true);
                //height = Utilities.PointToPixel(cellRect.Height, false);   

                #endregion

                left = (int)cellRect.Left;
                top = (int)cellRect.Top;
                width = (int)cellRect.Width;
                height = (int)cellRect.Height;
#if DEBUG
                System.Diagnostics.Debug.WriteLine(string.Format("Rectangle {0}", cellRect));
                System.Diagnostics.Debug.WriteLine(string.Format("Left {0} / Top {1} / Width {2} / Height {3}", left, top, width, height));     
#endif
            }
        }

        public override object GetPropertyValue(string propertyName)
        {
            switch (propertyName)
            {
                case PropertyNames.RowIndex:
                    return this.CellInfo.RowIndex;
                case PropertyNames.ColumnIndex:
                    return this.CellInfo.ColumnIndex;
                case PropertyNames.GridName:
                    return this.CellInfo.GridName;
            }

            return base.GetPropertyValue(propertyName);
        }

        public override IQueryElement QueryId
        {
            get
            {
                if (this.queryId == null)
                {
                    this.queryId = new QueryElement();
                    this.queryId.Condition = new AndCondition(
                            new PropertyCondition(PropertyNames.ControlType, this.ControlTypeName),
                            new PropertyCondition(PropertyNames.RowIndex, this.CellInfo.RowIndex),
                            new PropertyCondition(PropertyNames.ColumnIndex, this.CellInfo.ColumnIndex),
                            new PropertyCondition(PropertyNames.GridName,this.CellInfo.GridName));
                    this.queryId.Ancestor = this.Parent;
                }

                return this.queryId;
            }
        }

        #endregion

        public override void EnsureVisibleByScrolling(int pointX, int pointY, ref int outpointX, ref int outpointY)
        {
            // TODO - Add code to ScrollInView here
        }

        public override bool Equals(IUITechnologyElement element)
        {
            if (base.Equals(element))
            {
                GridCellElement otherElement = element as GridCellElement;
                if (otherElement != null)
                {
                    return this.CellInfo.Equals(otherElement.CellInfo);
                }
            }

            return false;
        }

        

        internal GridCellInfo CellInfo { get; private set; }

        internal GridCellElement(IntPtr windowHandle, GridCellInfo cellInfo, GridTechnologyManager manager)
            : base(windowHandle, manager)
        {
            this.CellInfo = cellInfo;
        }

        /// <summary>
        /// Gets the parent of this control in this technology.
        /// </summary>
        internal override UITechnologyElement Parent
        {
            get
            {
                if (this.parent == null)
                {
                    this.parent = this.technologyManager.GetGridElement(this.WindowHandle, this.CellInfo.Parent);
                }

                return this.parent;
            }
        }

        internal override System.Collections.IEnumerator GetChildren(AndCondition condition)
        {
            return null;
        }
    }
}
