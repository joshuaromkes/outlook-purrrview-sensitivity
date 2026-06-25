using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook_Purview_Sensitivity
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Debug.WriteLine("[PS] Startup firing");

            try
            {
                if (this.Application == null)
                {
                    Debug.WriteLine("[PS] ERROR: Application is null");
                    return;
                }

                Debug.WriteLine("[PS] Application OK");

                Outlook.Explorer explorer = this.Application.ActiveExplorer();
                if (explorer != null)
                {
                    Debug.WriteLine("[PS] Explorer found, wiring up");
                    WireUpExplorer(explorer);
                }
                else
                {
                    Debug.WriteLine("[PS] No explorer, hooking NewExplorer");
                    if (this.Application.Explorers != null)
                    {
                        this.Application.Explorers.NewExplorer += Explorers_NewExplorer;
                    }
                    else
                    {
                        Debug.WriteLine("[PS] ERROR: Explorers is null");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[PS] Startup error: {ex}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            Debug.WriteLine("[PS] NewExplorer fired");
            WireUpExplorer(explorer);
        }

        private void WireUpExplorer(Outlook.Explorer explorer)
        {
            Debug.WriteLine("[PS] WireUpExplorer");
            if (explorer == null) return;

            explorer.FolderSwitch += Explorer_FolderSwitch;

            Outlook.MAPIFolder folder = explorer.CurrentFolder;
            if (folder != null)
            {
                Debug.WriteLine($"[PS] CurrentFolder: {folder.Name}");
                ColumnManager.EnsureColumn(folder);
                ColumnManager.StampFolder(folder, maxItems: 50);
                Marshal.ReleaseComObject(folder);
            }
            else
            {
                Debug.WriteLine("[PS] ERROR: CurrentFolder is null");
            }
        }

        private void Explorer_FolderSwitch()
        {
            Debug.WriteLine("[PS] FolderSwitch fired");
            Outlook.Explorer explorer = this.Application?.ActiveExplorer();
            if (explorer == null) return;

            Outlook.MAPIFolder folder = explorer.CurrentFolder;
            if (folder == null) return;

            Debug.WriteLine($"[PS] Switched to: {folder.Name}");
            ColumnManager.EnsureColumn(folder);
            ColumnManager.StampFolder(folder, maxItems: 50);
            Marshal.ReleaseComObject(folder);
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
