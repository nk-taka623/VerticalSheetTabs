using System;
using System.IO;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace VerticalSheetTabs
{
    // Ribbon button "タブ表示の設定" on the View tab, and the same item in the menu of Excel's own
    // (horizontal) sheet tabs.
    [ComVisible(true)]
    public class RibbonController : Office.IRibbonExtensibility
    {
        private readonly Action _openSettings;

        public RibbonController(Action openSettings)
        {
            _openSettings = openSettings;
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabView"">
        <group id=""VTabGroup"" label=""縦タブ"">
          <button id=""VTabSettings"" label=""タブ表示の設定"" size=""large""
                  getImage=""GetSettingsImage""
                  screentip=""タブ表示の設定""
                  supertip=""シートタブの表示モード (横のみ/縦のみ/両方) と、縦タブの位置 (左/右) を設定します。""
                  onAction=""OnSettingsClick"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuWorkbookPly"">
      <button id=""VTabSettingsTabMenu"" label=""タブ表示の設定""
              insertBeforeMso=""SheetInsertPage""
              onAction=""OnSettingsClick"" />
      <menuSeparator id=""VTabSettingsTabMenuSeparator"" insertBeforeMso=""SheetInsertPage"" />
    </contextMenu>
  </contextMenus>
</customUI>";
        }

        // The button's own icon (see RibbonIcon). It is drawn large and Office scales it to the DPI.
        public System.Drawing.Bitmap GetSettingsImage(Office.IRibbonControl control)
        {
            return RibbonIcon.Create(64);
        }

        public void OnSettingsClick(Office.IRibbonControl control)
        {
            _openSettings();
        }
    }
}
