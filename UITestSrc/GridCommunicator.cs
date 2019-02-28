using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Syncfusion.UITest.GridCommunication;

namespace Syncfusion.Grid.WPF.UITest
{
    public class GridCommunicator
    {
        private static IGridInteropService thisInstance;
        public static IGridInteropService Instance
        {
            get
            {
                if (thisInstance == null)
                {
                    thisInstance = (IGridInteropService)Activator.GetObject(typeof(IGridInteropService), "ipc://GridControl/GridTestService");
                }

                return thisInstance;
            }
        }
    }
}
