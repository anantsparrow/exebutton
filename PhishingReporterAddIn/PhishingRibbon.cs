using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace PhishingReporterAddIn
{
    [ComVisible(true)]
    public class PhishingRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public PhishingRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PhishingReporterAddIn.PhishingRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnReportPhishing(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.Instance.ReportPhishingEmail();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An unexpected error occurred during execution: " + ex.Message, "Phishing Triage Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public System.Drawing.Bitmap GetImage(string imageName)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                // Load from embedded resource stream
                using (Stream stream = asm.GetManifestResourceStream("PhishingReporterAddIn.Resources." + imageName + ".png"))
                {
                    if (stream != null)
                    {
                        return new System.Drawing.Bitmap(stream);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error loading custom ribbon image: " + ex.Message);
            }
            return null;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] names = asm.GetManifestResourceNames();
            for (int i = 0; i < names.Length; i++)
            {
                if (string.Compare(resourceName, names[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(names[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
