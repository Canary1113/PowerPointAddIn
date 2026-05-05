using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using PowerPointAddIn.Effects;

namespace PowerPointAddIn.UI
{
    public partial class PPTAssistantRibbon
    {
        private Image themeIcon;
        private readonly Dictionary<RibbonDropDownItem, ThemeTemplateInfo> themeTemplates =
            new Dictionary<RibbonDropDownItem, ThemeTemplateInfo>();

        private void PPTAssistantRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            themeIcon = CreateThemeIcon();
            galleryTheme.OfficeImageId = string.Empty;
            galleryTheme.Image = themeIcon;
        }

        private void buttonApplyLiquidGlass_Click(object sender, RibbonControlEventArgs e)
        {
            Execute(debugMode: false);
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

        private void galleryTheme_Click(object sender, RibbonControlEventArgs e)
        {
            var selectedItem = galleryTheme.SelectedItem;

            if (selectedItem != null && themeTemplates.TryGetValue(selectedItem, out ThemeTemplateInfo template))
            {
                ExecuteTheme(service => service.ApplyTemplate(template));
            }
            galleryTheme.SelectedItemIndex = -1;
        }

        private void galleryTheme_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            RefreshThemeGallery(showErrors: true);
        }

        private void buttonThemeColor_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteThemePersonalize();
        }

        private void buttonOpenThemeTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteTheme(service => service.OpenThemeTemplate());
        }

        private static void ExecuteThemePersonalize()
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = Color.FromArgb(0, 112, 192);
                colorDialog.FullOpen = true;

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    ExecuteTheme(service => service.CustomizeAccentColor(colorDialog.Color));
                }
            }
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

        private void RefreshThemeGallery(bool showErrors)
        {
            themeTemplates.Clear();
            galleryTheme.Items.Clear();

            if (ThisAddIn.Instance == null)
            {
                return;
            }

            try
            {
                var service = new ThemeService(ThisAddIn.Instance.Application);
                foreach (ThemeTemplateInfo template in service.GetAvailableTemplates())
                {
                    RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                    item.Label = template.Label;
                    galleryTheme.Items.Add(item);
                    themeTemplates[item] = template;
                }
            }
            catch (Exception ex)
            {
                if (showErrors)
                {
                    MessageBox.Show(
                        ex.Message,
                        "PPT Assistant",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
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

        private static Image CreateThemeIcon()
        {
            var bitmap = new Bitmap(32, 32);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var outlinePen = new Pen(Color.FromArgb(80, 80, 80), 1.4f))
            using (var accentPen = new Pen(Color.FromArgb(0, 112, 192), 2f))
            using (var whiteBrush = new SolidBrush(Color.White))
            using (var lightBrush = new SolidBrush(Color.FromArgb(247, 247, 247)))
            using (var darkBrush = new SolidBrush(Color.FromArgb(64, 64, 64)))
            using (var accentBrush = new SolidBrush(Color.FromArgb(0, 112, 192)))
            {
                graphics.Clear(Color.Transparent);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                graphics.FillRectangle(lightBrush, 5, 6, 21, 17);
                graphics.DrawRectangle(outlinePen, 5, 6, 21, 17);
                graphics.FillRectangle(whiteBrush, 8, 9, 15, 3);
                graphics.FillRectangle(darkBrush, 8, 15, 8, 2);
                graphics.FillRectangle(darkBrush, 8, 19, 12, 2);

                graphics.FillEllipse(accentBrush, 19, 17, 10, 10);
                graphics.DrawEllipse(accentPen, 19, 17, 10, 10);
                graphics.DrawLine(accentPen, 22, 22, 26, 22);
                graphics.DrawLine(accentPen, 24, 20, 24, 24);
            }

            return bitmap;
        }
    }
}
