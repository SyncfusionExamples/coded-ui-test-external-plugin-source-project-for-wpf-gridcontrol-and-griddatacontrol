using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows;

namespace Syncfusion.UITest.GridCommunication
{
    /// <summary>
    /// Interface for get/set Grid cell info using IPC 
    /// </summary>
    public interface IGridInteropService
    {
        /// <summary>
        /// Get property value from cell.
        /// </summary>
        /// <param name="style"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        object GetStyleInfoProperty(GridCellInfo style, string propertyName);

        /// <summary>
        /// Set style info property to cell
        /// </summary>
        /// <param name="style"></param>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        void SetStyleInfoProperty(GridCellInfo style, string propertyName, object propertyValue);

        /// <summary>
        /// Get rect value of corresponding CellInfo
        /// </summary>
        /// <param name="cellInfo"></param>
        /// <returns></returns>
        Rect GetCellBoundingRect(GridCellInfo cellInfo);

        /// <summary>
        /// Get Element from the corresponding point in Grid
        /// </summary>
        /// <param name="pointX"></param>
        /// <param name="pointY"></param>
        /// <returns></returns>
        GridCellInfo GetElementFromPoint(int pointX, int pointY);
    }

    [Serializable]
    public abstract class GridInfo
    {
    }

    [Serializable]
    public class GridControlInfo : GridInfo
    {
    }

    [Serializable]
    public class GridCellInfo : GridInfo
    {
        private static GridCellInfo empty;
        public static GridCellInfo Empty
        {
            get
            {
                if (empty == null)
                {
                    empty = new GridCellInfo(int.MaxValue, int.MaxValue, "");
                }

                return empty;
            }
        }

        public GridCellInfo(int rowIndex, int colIndex, string gridName)
        {
            this.RowIndex = rowIndex;
            this.ColumnIndex = colIndex;
            this.GridName = gridName;
        }

        /// <summary>
        /// RowIndex of the Cell
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// ColumnIndex of the Cell
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// GridName of the cell
        /// </summary>
        public string GridName { get; set; }

        /// <summary>
        ///  Helpful in debugging.
        /// </summary>
        /// <returns>cell info as string</returns>
        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "[{0}, {1}],{2}", RowIndex, ColumnIndex, GridName);
        }

        /// <summary>
        /// Parent of Cell
        /// </summary>
        public GridControlInfo Parent { get; set; }
    }
}
