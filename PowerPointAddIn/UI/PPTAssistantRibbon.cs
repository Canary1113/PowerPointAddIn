using System;
using System.Drawing;
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

        private void buttonGlassWhite_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteLayeredGlass(Color.White);
        }

        private void buttonGlassBlack_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteLayeredGlass(Color.Black);
        }

        private void buttonRecover_Click(object sender, RibbonControlEventArgs e)
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
                service.RecoverSelection(debugMode: false);
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

        private void buttonNavbarAnimation_Click(object sender, RibbonControlEventArgs e)
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
                var service = new NavbarAnimationService(ThisAddIn.Instance.Application);
                service.ApplyToSelection();
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

        private void buttonThemeStyle1_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteTheme(service => service.ApplyStyle1());
        }

        private void buttonThemeStyle2_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteTheme(service => service.ApplyStyle2());
        }

        private static void ExecuteTheme(Action<ThemeService> applyTheme)
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
                var service = new ThemeService(ThisAddIn.Instance.Application);
                applyTheme(service);
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

        private static void ExecuteLayeredGlass(Color overlayColor)
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
                service.ApplyLayeredToSelection(debugMode: false, overlayColor);
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
