using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITest.Common;
using UITestClass = Microsoft.VisualStudio.TestTools.UITest.Common.UITest;
using Microsoft.VisualStudio.TestTools.UITest.Common.UIMap;

[assembly: UITestExtensionPackage("GridControlUITestExtensionPackage", typeof(Syncfusion.Grid.WPF.UITest.GridControlUITestExtensionPackage))]

namespace Syncfusion.Grid.WPF.UITest
{
    internal class GridControlUITestExtensionPackage : UITestExtensionPackage
    {
        private bool savingHandlerAttached;
        private GridTechnologyManager technologyManager = new GridTechnologyManager();
        private UITestPropertyProvider[] propertyProviders = new UITestPropertyProvider[] { new GridPropertyProvider(), new CommonPropertyProvider() };

        #region Public method

        public override void Dispose()
        {
            // Detach handler
            if (savingHandlerAttached)
            {
                savingHandlerAttached = false;
                UITestClass.Saving -= new EventHandler<UITestEventArgs>(UITestSaving);
            }
        }

        public override object GetService(Type serviceType)
        {
            // Attach handler if not attached
            if (!savingHandlerAttached)
            {
                savingHandlerAttached = true;
                UITestClass.Saving += new EventHandler<UITestEventArgs>(UITestSaving);
            }

            if (serviceType == typeof(UITechnologyManager))
            {
                return this.technologyManager;
            }
            else if (serviceType == typeof(UITestPropertyProvider))
            {
                return this.propertyProviders;
            }

            return null;
        }

        #endregion

        #region private method

        /// <summary>
        /// change all the elements to UIA elements except GridCell.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UITestSaving(object sender, UITestEventArgs e)
        {
            if (e.UITest != null && e.UITest.Maps != null && e.UITest.Maps.Count == 1)
            {
                // At this point, inspect all the UIObject recursively for any GridControl element.
                foreach (var topLevelElement in e.UITest.Maps[0].TopLevelWindows)
                {
                    FixTechnologyManager(topLevelElement);
                }
            }
        }

        /// <summary>
        /// Recursively fixes the technology manager for all the controls to except Control type is Cell.
        /// </summary>
        /// <param name="uiObject">The UI control.</param>
        private void FixTechnologyManager(UIObject uiObject)
        {
            if (uiObject.ControlType == "Cell")
                return;
            // If technology name is GridControl, change it to UIA
            if (string.Equals(uiObject.TechnologyName, Utilities.GridControlTechnologyName, StringComparison.OrdinalIgnoreCase))
            {
                uiObject.TechnologyName = "UIA";
            }

            // Recurse all descendants and do the same for those too.
            ICollection<UIObject> descendants = uiObject.Descendants;
            if (descendants != null)
            {
                foreach (var descendant in descendants)
                {
                    FixTechnologyManager(descendant);
                }
            }
        }

        #endregion

        #region Public properties

        public override string PackageDescription
        {
            get { return "Plugin for VSTT Reocrd and Playback support on GridControl applications"; }
        }

        public override string PackageName
        {
            get { return "VSTT GridControl Plugin"; }
        }

        public override string PackageVendor
        {
            get { return "Syncfusion Inc.,"; }
        }

        public override Version PackageVersion
        {
            get { return new Version(1, 0); }
        }

        public override Version VSVersion
        {
            get { return new Version(10, 0); }
        }

        #endregion
    }
}
