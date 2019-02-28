using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Syncfusion.Windows.Controls.Grid;
using System.Windows;
using System.Diagnostics;
using Syncfusion.Windows.Controls.Cells;
using Syncfusion.Windows.ComponentModel;
using Syncfusion.Windows.Collections;
using System.ComponentModel;
using System.Windows.Media;
using System.Windows.Input;

namespace Syncfusion.UITest.GridCommunication
{
    public class GridInteropService : MarshalByRefObject, IGridInteropService
    {
        private GridControlBase grid;
        private GridControlBase defaultGrid;
        private Window window;
        private Dictionary<GridControlBase, string> GridWithNames = new Dictionary<GridControlBase, string>();

        #region contructor

        public GridInteropService()
        {
            GridControlTestApplication.Current.Dispatcher.BeginInvoke(new Action(() =>
            {
                //in realtime application (like using MVVM pattern) GridControlTestApplication.Current.MainWindow value is null.
                //cant get the window from GridControlTestApplication.Current.MainWindow.instead of this
                //we use GridControlTestApplication.Current.Windows.OfType<Window>() to get the windows.
                window = GridControlTestApplication.Current.Windows.OfType<Window>().FirstOrDefault();
                if (window != null)
                {
                    defaultGrid = window.FindElementOfType<GridControlBase>();
                    var GridLists = window.FindElementsOfType<GridControlBase>();
                    foreach (GridControlBase grid in GridLists)
                        if (grid != null && (grid.GetType() == typeof(GridDataControlBaseImpl) || grid.GetType() == typeof(GridControl) || grid.GetType() == typeof(GridTreeControlImpl)))
                        {
                            //WPF-28663 if grid does nto have name then we assign some name for identify uniqly in automation
                            if (string.IsNullOrEmpty(grid.Name))
                                grid.Name = "Grid" + (GridWithNames.Count + 1);
                            GridWithNames.Add(grid, grid.Name);
                        }
                }
            }));
        }

        #endregion

        #region Public methods
        /// <summary>
        /// Invokes to get value for Particular property
        /// </summary>
        /// <param name="cellInfo">GridCellInfo</param>
        /// <param name="propertyName">propertyName</param>
        /// <returns> value of particaular property</returns>
        public object GetStyleInfoProperty(GridCellInfo cellInfo, string propertyName)
        {
            var style = this.GetStyle(cellInfo);
            switch (propertyName)
            {
                case PropertyNames.Text:
                    {
                        var value = style.CellValue == null ? string.Empty : style.CellValue;
                        var valueType = value.GetType();
                        return valueType.IsSerializable ? style.Text : "Non Serializable value";
                    }
                case PropertyNames.CellValue:
                    {
                        var value = style.CellValue == null ? string.Empty : style.CellValue;
                        var valueType = value.GetType();
                        var objectValue = this.GetObjectValue(value);
                        return objectValue;
                    }
                case PropertyNames.ColumnHeader:
                    {
                        string headertext = string.Empty;
                        var info = this.grid.Model[cellInfo.RowIndex, cellInfo.ColumnIndex];
                        if (info != null && info is GridDataStyleInfo)
                        {
                            var styleInfo=(info as GridDataStyleInfo);
                            if (styleInfo.CellIdentity.Column != null)
                                this.window.Dispatcher.Invoke((Action)(() =>
                                {
                                    headertext = styleInfo.CellIdentity.Column.HeaderText;
                                }));
                        }
                        return headertext;
                    }
                case PropertyNames.Description:
                    return style.HasDescription ? style.Description : string.Empty;
                case PropertyNames.SelectedRanges:
                    return this.grid.Model.SelectedRanges.ToString();
                case PropertyNames.Format:
                    return style.Format != null ? style.Format : string.Empty;
                case PropertyNames.FormulaTag:
                    return style.FormulaTag != null ? style.FormulaTag.ToString() : string.Empty;
                case PropertyNames.CellWidth:
                    {
                        var width = this.grid.Model.ColumnWidths[cellInfo.ColumnIndex];
                        return width.ToString();
                    }
                case PropertyNames.CellHeight:
                    {
                        var height = this.grid.Model.RowHeights[cellInfo.RowIndex];
                        return height.ToString();
                    }
                case PropertyNames.FormattedText:
                    return style.FormattedText != null ? style.FormattedText : string.Empty;
                case PropertyNames.RowCount:
                    return this.grid.Model.RowCount;
                case PropertyNames.ColumnCount:
                    return this.grid.Model.ColumnCount;
                case PropertyNames.GridName:
                    if (GridWithNames.ContainsKey(this.grid))
                        return GridWithNames[this.grid];
                    else
                        return string.Empty;
                default:
                    throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Invokes to set the property value
        /// </summary>
        /// <param name="cellInfo">GridCellInfo</param>
        /// <param name="propertyName">propertyName</param>
        /// <param name="propertyValue">propertyValue</param>
        public void SetStyleInfoProperty(GridCellInfo cellInfo, string propertyName, object propertyValue)
        {
            var style = this.GetStyle(cellInfo);
            switch (propertyName)
            {
                case PropertyNames.CellValue:
                    style.CellValue = propertyValue;
                    break;
            }
        }

        /// <summary>
        /// Invoke to get cell bound rect value
        /// </summary>
        /// <param name="cellInfo">GridCellInfo</param>
        /// <returns>Rect</returns>
        public Rect GetCellBoundingRect(GridCellInfo cellInfo)
        {
            var currentCellRange = GridRangeInfo.Cell(cellInfo.RowIndex, cellInfo.ColumnIndex);
            this.grid = GetGrid(cellInfo.GridName);
            var cc = this.grid.CoveredCells.GetCellSpan(cellInfo.RowIndex, cellInfo.ColumnIndex);
            if (cc != null)
            {
                currentCellRange = GridRangeInfo.Cells(cc.Top, cc.Left, cc.Bottom, cc.Right);
            }

            var yCurrentCellPos = this.grid.ScrollRows.RangeToRegionPoints(currentCellRange.Top, currentCellRange.Bottom, true);
            var xCurrentCellPos = this.grid.ScrollColumns.RangeToRegionPoints(currentCellRange.Left, currentCellRange.Right, true);
            var rowRegion = 1;
            var colRegion = 1;
            if (yCurrentCellPos[rowRegion].IsEmpty)
            {
                return Rect.Empty;
            }

            if (xCurrentCellPos[colRegion].IsEmpty)
            {
                return Rect.Empty;
            }

            var point = new Point(xCurrentCellPos[colRegion].Start, yCurrentCellPos[rowRegion].Start);
            GridControlTestApplication.Current.Dispatcher.Invoke(new Action(() =>
            {
                point = this.grid.PointToScreen(point);
            }), System.Windows.Threading.DispatcherPriority.Send);
            var rect = new Rect(point.X, point.Y, xCurrentCellPos[colRegion].Length, yCurrentCellPos[rowRegion].Length);
            return rect;
        }

        /// <summary>
        /// Get the GridCellInfo from particaulr point.if it is not a cell then retrun Null.
        /// </summary>
        /// <param name="pointX">x position</param>
        /// <param name="pointY">y position</param>
        /// <returns>GridCellInfo</returns>
        public GridCellInfo GetElementFromPoint(int pointX, int pointY)
        {
            Point point = new Point(pointX, pointY);
            if (this.window == null)
                return null;
            this.window.Dispatcher.CheckAccess();

            #region cell detection Area

            bool isPointConverted = false;
            GridCellInfo cell = null;
            this.window.Dispatcher.Invoke((Action)(() =>
             {
                 DependencyObject dt = Mouse.DirectlyOver as DependencyObject;
                 if (dt != null && dt is GridControlBase)
                 {
                     this.grid = dt as GridControlBase;
                     if (!GridWithNames.ContainsKey(grid))
                     {
                         //WPF-28663 if grid does not have name then we assign some name for identify uniqe in automation
                         if (string.IsNullOrEmpty(grid.Name))
                             grid.Name = "Grid" + (GridWithNames.Count + 1);
                         GridWithNames.Add(grid, grid.Name);
                     }
                     //WPF-28067 when large application with multiple windows, while navigate from one window to another window 
                     //while cross hair is running "The visual is not connected to a PresentationSource" exception is throws.
                     //so need to put PointFromScreen method inside the try catch block to avoid exception
                     try
                     {
                         point = this.grid.PointFromScreen(point);
                     }
                     catch
                     {
                         //not need to handle it
                     }
                     isPointConverted = true;
                 }
                 else
                 {
                     foreach (GridControlBase grid in GridWithNames.Keys)
                     {
                         var gridPt = grid.PointToScreen(new Point());
                         var gridRect = new Rect(gridPt, grid.DesiredSize);
                         if (gridRect.Contains(new Point(pointX, pointY)))
                         {
                             this.grid = grid;
                         }
                     }
                 }

             }), System.Windows.Threading.DispatcherPriority.Send);

            if (this.grid != null)
            {
                if (!isPointConverted)
                    this.window.Dispatcher.Invoke((Action)(() =>
                            {
                                try
                                {
                                    point = this.grid.PointFromScreen(point);
                                }
                                catch
                                {
                                    //not need to handle it
                                }
                            }), System.Windows.Threading.DispatcherPriority.Send);

                var rowcol = this.grid.PointToCellRowColumnIndexOutsideCells(point, false);

                if (rowcol.IsEmpty || rowcol.ColumnIndex == -1 || rowcol.RowIndex == -1)
                {
                    return null;
                }

                cell = new GridCellInfo(rowcol.RowIndex, rowcol.ColumnIndex, GridWithNames.FirstOrDefault(x => x.Key == this.grid).Value);
            }

            return cell;

            #endregion
        }

        #endregion

        #region Privae Methods

        private object GetObjectValue(object value)
        {
            var type = value.GetType();
            if (!type.IsSerializable)
            {
                return "Non Serializable Value";
            }

            if (ListUtil.IsComplexType(type))
            {
                var sb = new StringBuilder();
                var pDesc = TypeDescriptor.GetProperties(value);
                foreach (PropertyDescriptor pd in pDesc)
                {
                    sb.Append(string.Format("{0} : {1}, ", pd.Name, pd.GetValue(value)));
                }
                return sb.ToString();
            }
            //WPF-28414 if cell value is enum type then we must convert it to string.
            if(type.IsEnum)
            {
                return value.ToString();
            }

            return value;
        }

        private GridStyleInfo GetStyle(GridCellInfo cellInfo)
        {
            var currentCellRowIndex = new RowColumnIndex(cellInfo.RowIndex, cellInfo.ColumnIndex);
            this.grid = GetGrid(cellInfo.GridName);
            var cc = this.grid.CoveredCells.GetCellSpan(cellInfo.RowIndex, cellInfo.ColumnIndex);
            if (cc != null)
            {
                currentCellRowIndex = new RowColumnIndex(cellInfo.RowIndex, cc.Left);
            }
            return this.grid.GetRenderStyleInfo(currentCellRowIndex.RowIndex, currentCellRowIndex.ColumnIndex);
        }

        private GridControlBase GetGrid(string gridName)
        {
            if (!string.IsNullOrEmpty(gridName) && GridWithNames.ContainsValue(gridName))
                return GridWithNames.FirstOrDefault(x => x.Value == gridName).Key;

            if (defaultGrid == null)
                throw new NullReferenceException();
            return defaultGrid;
        }

        #endregion
    }
}
