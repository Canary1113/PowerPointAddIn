using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPointAddIn.Effects;

namespace PowerPointAddIn.UI
{
    internal sealed class LiquidGlassRibbon : IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonId)
        {
            MessageBox.Show(
                $"Ribbon XML requested for: {ribbonId}",
                "PPT Assistant",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return GetResourceText("PowerPointAddIn.UI.LiquidGlassRibbon.xml");
        }

        public void OnApplyLiquidGlass(IRibbonControl control)
        {
            Execute(debugMode: false);
        }

        public void OnApplyLiquidGlassDebug(IRibbonControl control)
        {
            Execute(debugMode: true);
        }

        private static string GetResourceText(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException("Ribbon XML resource was not found.");
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static void Execute(bool debugMode)
        {
            if (ThisAddIn.Instance == null)
            {
                MessageBox.Show(
                    "The add-in is not initialized yet.",
                    "PPT Assistant",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var service = new LiquidGlassEffectService(ThisAddIn.Instance.Application);
                int changedCount = service.ApplyToSelection(debugMode);

                if (debugMode)
                {
                    MessageBox.Show(
                        $"Finished. {changedCount} shape(s) were updated.",
                        "PPT Assistant",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "PPT Assistant",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
