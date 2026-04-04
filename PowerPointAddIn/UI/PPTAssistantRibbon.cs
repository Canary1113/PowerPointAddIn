using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using PowerPointAddIn.Effects;

namespace PowerPointAddIn.UI
{
    public partial class PPTAssistantRibbon
    {
        private void PPTAssistantRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonApplyLiquidGlass_Click(object sender, RibbonControlEventArgs e)
        {
            Execute(debugMode: false);
        }

        private void buttonApplyLiquidGlassDebug_Click(object sender, RibbonControlEventArgs e)
        {
            Execute(debugMode: true);
        }

        private void buttonBackground_Click(object sender, RibbonControlEventArgs e)
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
                var service = new BackgroundImageService(ThisAddIn.Instance.Application);
                service.ApplyFromSelectedPicture();
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
