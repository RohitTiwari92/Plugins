using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PluginPowerPoint
{
    public partial class ThisAddIn
    {
        private BackPane panebase;
        private static Microsoft.Office.Tools.CustomTaskPane TaskPaneObj;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            panebase = new BackPane();
            TaskPaneObj = this.CustomTaskPanes.Add(panebase, "SmartArt Count");
            TaskPaneObj.Visible = false;
           // TaskPaneObj.VisibleChanged += TaskPane_VisibleChanged; 
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return TaskPaneObj;
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
