using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookPurviewColumn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                var explorer = this.Application.ActiveExplorer();
                if (explorer != null)
                {
                    WireUpExplorer(explorer);
                }
                else
                {
                    // No explorer window on startup (rare but possible).
                    // Hook NewExplorer to catch the first window that opens.
                    this.Application.Explorers.NewExplorer += Explorers_NewExplorer;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[OutlookPurviewColumn] Startup error: {ex}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            WireUpExplorer(explorer);
        }

        private void WireUpExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null) return;
            explorer.FolderSwitch += Explorer_FolderSwitch;

            var folder = explorer.CurrentFolder;
            if (folder != null)
            {
                ColumnManager.EnsureColumn(folder);
                ColumnManager.StampFolder(folder, maxItems: 50);
                Marshal.ReleaseComObject(folder);
            }
        }

        private void Explorer_FolderSwitch()
        {
            var explorer = this.Application.ActiveExplorer();
            if (explorer == null) return;

            var folder = explorer.CurrentFolder;
            if (folder != null)
            {
                ColumnManager.EnsureColumn(folder);
                ColumnManager.StampFolder(folder, maxItems: 50);
                Marshal.ReleaseComObject(folder);
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
