using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class LiquidGlassEffectService
    {
        private readonly PowerPoint.Application application;

        public LiquidGlassEffectService(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public int ApplyToSelection(bool debugMode)
        {
            PowerPoint.DocumentWindow activeWindow = application.ActiveWindow;
            if (activeWindow == null)
            {
                throw new InvalidOperationException("No active PowerPoint window was found.");
            }

            PowerPoint.Selection selection = activeWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                throw new InvalidOperationException("Select one or more shapes first.");
            }

            PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
            int changedCount = 0;

            for (int index = 1; index <= shapeRange.Count; index++)
            {
                changedCount += ApplyToShapeRecursive(shapeRange[index], debugMode);
            }

            if (changedCount == 0)
            {
                throw new InvalidOperationException("No supported shapes were found in the current selection.");
            }

            return changedCount;
        }

        private int ApplyToShapeRecursive(PowerPoint.Shape shape, bool debugMode)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                int groupCount = 0;
                PowerPoint.GroupShapes groupItems = shape.GroupItems;

                for (int index = 1; index <= groupItems.Count; index++)
                {
                    groupCount += ApplyToShapeRecursive(groupItems[index], debugMode);
                }

                return groupCount;
            }

            if (!SupportsLiquidGlass(shape))
            {
                return 0;
            }

            DebugStep(debugMode, $"Applying liquid-glass effect to: {shape.Name}");

            ApplyBackgroundFill(shape);
            ApplyHighlightOutline(shape);
            ApplyShadow(shape);

            return 1;
        }

        private static bool SupportsLiquidGlass(PowerPoint.Shape shape)
        {
            if (shape.Width < 6f || shape.Height < 6f)
            {
                return false;
            }

            switch (shape.Type)
            {
                case MsoShapeType.msoLine:
                case MsoShapeType.msoPicture:
                case MsoShapeType.msoLinkedPicture:
                case MsoShapeType.msoMedia:
                case MsoShapeType.msoEmbeddedOLEObject:
                case MsoShapeType.msoLinkedOLEObject:
                case MsoShapeType.msoChart:
                case MsoShapeType.msoTable:
                case MsoShapeType.msoSmartArt:
                case MsoShapeType.msoCanvas:
                    return false;
                default:
                    return true;
            }
        }

        private static void ApplyBackgroundFill(PowerPoint.Shape shape)
        {
            shape.Fill.Visible = MsoTriState.msoTrue;
            shape.Fill.Background();
        }

        private static void ApplyHighlightOutline(PowerPoint.Shape shape)
        {
            float baseSize = Math.Min(shape.Width, shape.Height);
            float lineWeight = Clamp(baseSize * 0.018f, 1.25f, 4.5f);

            shape.Line.Visible = MsoTriState.msoTrue;
            shape.Line.Style = MsoLineStyle.msoLineSingle;
            shape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            shape.Line.InsetPen = MsoTriState.msoTrue;
            shape.Line.Weight = lineWeight;
            shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(255, 255, 255));
            shape.Line.Transparency = 0.28f;

            shape.Glow.Radius = Clamp(lineWeight * 2.1f, 2.5f, 10f);
            shape.Glow.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(255, 255, 255));
            shape.Glow.Transparency = 0.52f;
        }

        private static void ApplyShadow(PowerPoint.Shape shape)
        {
            float baseSize = Math.Min(shape.Width, shape.Height);
            float blur = Clamp(baseSize * 0.03f, 2f, 12f);
            float offset = Clamp(baseSize * 0.008f, 0.75f, 4f);

            shape.Shadow.Visible = MsoTriState.msoTrue;
            shape.Shadow.Blur = blur;
            shape.Shadow.OffsetX = 0f;
            shape.Shadow.OffsetY = offset;
            shape.Shadow.Transparency = 0.82f;
            shape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(176, 210, 235));
        }

        private static float Clamp(float value, float min, float max)
        {
            return Math.Max(min, Math.Min(max, value));
        }

        private static void DebugStep(bool enabled, string message)
        {
            if (!enabled)
            {
                return;
            }

            MessageBox.Show(
                message,
                "Apply Glass Debug",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
