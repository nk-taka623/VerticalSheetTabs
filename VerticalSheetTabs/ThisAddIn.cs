using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace VerticalSheetTabs
{
    public partial class ThisAddIn
    {
        private PaneManager _panes;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _panes = new PaneManager(this.CustomTaskPanes);

            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;

            this.Application.SheetActivate += sh => _panes.SyncAll();

            foreach (Excel.Workbook wb in this.Application.Workbooks)
            {
                foreach (Excel.Window w in wb.Windows)
                {
                    _panes.EnsurePane(w, wb);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_WindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            _panes.EnsurePane(wn, wb);
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            foreach (Excel.Window w in wb.Windows)
            {
                _panes.RemovePane(w);
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
