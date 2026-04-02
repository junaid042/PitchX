using System;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;

namespace PitchX
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Check for updates silently in background
            Task.Run(async () =>
            {
                await Task.Delay(5000); // Wait 5s for PowerPoint to finish loading
                await TLBR.AutoUpdater.CheckForUpdatesAsync();
            });
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        // IMPORTANT: return the XML-based ribbon object here
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new TLBRRibbon();
        }
    }
}
