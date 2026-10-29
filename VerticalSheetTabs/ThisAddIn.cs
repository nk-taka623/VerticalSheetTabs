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

        // An exception that escapes an event handler would show Excel's error dialog or, worse,
        // stop the add-in. Every handler is therefore guarded: the next timer tick or event retries.
        private static void Guard(System.Action action)
        {
            try
            {
                action();
            }
            catch (Exception)
            {
                // Excel is busy, a window or workbook is closing, or the add-in is shutting down.
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Guard(() =>
            {
                DpiHelper.ExcelWindow = new IntPtr(this.Application.Hwnd);
                _panes = new PaneManager(this.CustomTaskPanes, this.Application);

                this.Application.WindowActivate += Application_WindowActivate;
                this.Application.SheetActivate += sh => Guard(() => _panes.SyncSheetOf(sh));
                // the badges of the rows (errors, notes): the sheet is looked at again when it has been changed or calculated, or left;
                // the events are listened to only while some window shows a badge
                _panes.BadgeEventsWanted = SetBadgeEvents;
                this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;

                _panes.EnsureAllPanes();
            });
        }

        private bool _badgeEvents;

        private void SetBadgeEvents(bool on)
        {
            if (on == _badgeEvents) return;
            _badgeEvents = on;
            if (on)
            {
                this.Application.SheetChange += Application_SheetChange;
                this.Application.SheetCalculate += Application_SheetCalculate;
                this.Application.SheetDeactivate += Application_SheetDeactivate;
            }
            else
            {
                this.Application.SheetChange -= Application_SheetChange;
                this.Application.SheetCalculate -= Application_SheetCalculate;
                this.Application.SheetDeactivate -= Application_SheetDeactivate;
            }
        }

        private void Application_SheetChange(object sh, Excel.Range target)
        {
            Guard(() => _panes.MarkBadgesDirty(sh));
        }

        private void Application_SheetCalculate(object sh)
        {
            Guard(() => _panes.MarkBadgesDirty(sh));
        }

        private void Application_SheetDeactivate(object sh)
        {
            Guard(() => _panes.MarkBadgesDirty(sh, true));
        }

        // Also runs when the user disables the add-in while Excel is running: the tab bar of
        // Excel is given back, so that no workbook is left without tabs.
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Guard(() => { if (_panes != null) { _panes.BadgeEventsWanted = null; SetBadgeEvents(false); _panes.Shutdown(); } });
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonController(() => { if (_panes != null) _panes.ShowSettings(); });
        }

        private void Application_WindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            Guard(() => _panes.EnsurePane(wn, wb));
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook wb, bool saveAsUi, ref bool cancel)
        {
            try
            {
                _panes.SaveSettings(wb);
            }
            catch (Exception)
            {
                // Never get in the way of the user's save.
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
