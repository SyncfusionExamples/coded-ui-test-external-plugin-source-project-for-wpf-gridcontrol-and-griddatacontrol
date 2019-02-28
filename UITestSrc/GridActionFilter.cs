using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UITest.Common;
using System.Windows.Forms;

namespace Syncfusion.Grid.WPF.UITest
{
    public class GridActionFilter : UITestActionFilter
    {
        /// <summary>
        /// Gets whether to apply aggregator timeout limits.
        /// </summary>
        /// <value></value>
        public override bool ApplyTimeout
        {
            get { return true; }
        }

        public override UITestActionFilterCategory Category
        {
            get { return UITestActionFilterCategory.PostSimpleToCompoundActionConversion; }
        }

        public override bool Enabled
        {
            get { return true; }
        }

        public override UITestActionFilterType FilterType
        {
            get { return UITestActionFilterType.Binary; }
        }

        public override string Group
        {
            get { return "GridControlActionFilters"; }
        }

        public override string Name
        {
            get { return "GridControlActionFilter"; }
        }

        public override bool ProcessRule(IUITestActionStack actionStack)
        {
            // Remove the mouse click preceeding on send keys on the same Excel cell.
            SendKeysAction lastAction = actionStack.Peek() as SendKeysAction;
            MouseAction secondLastAction = actionStack.Peek(1) as MouseAction;

            if (IsLeftClick(secondLastAction) &&
                AreActionsOnSameExcelCell(lastAction, secondLastAction))
            {
                // This is left click on a cell preceding a typing on the same cell.
                // Remove the left click action.
                // (0 means top-most action and 1 means 2nd action & so on.) 
                actionStack.Pop(1);
            }

            return false;
        }

        // Checks if this is a left click action or not.
        private static bool IsLeftClick(MouseAction mouseAction)
        {
            return mouseAction != null &&
                   mouseAction.ActionType == MouseActionType.Click &&
                   mouseAction.MouseButton == MouseButtons.Left;
                  // mouseAction.ModifierKeys == System.Windows.Input.ModifierKeys.None;
        }

        // Checks if two actions are on same cell or not.
        private static bool AreActionsOnSameExcelCell(SendKeysAction lastAction, MouseAction secondLastAction)
        {
            return lastAction != null && secondLastAction != null &&
                   lastAction.UIElement is GridCellElement &&
                   secondLastAction.UIElement is GridCellElement &&
                   object.Equals(lastAction.UIElement, secondLastAction.UIElement);
        }
    }
}
