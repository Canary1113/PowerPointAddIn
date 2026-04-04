using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class LiquidGlassEffectService
    {
        private const string GlassHelperTag = "PPTAssistantGlassHelper";
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

            RemoveLegacyHelperOutlines(shape);
            ApplyBackgroundFill(shape);
            ApplyHighlightOutline(shape);
            ApplyBevel(shape);
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
            shape.Line.Visible = MsoTriState.msoTrue;
            shape.Line.Style = MsoLineStyle.msoLineSingle;
            shape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            shape.Line.InsetPen = MsoTriState.msoTrue;
            shape.Line.Weight = 0.5f;
            shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(255, 255, 255));
            shape.Line.BackColor.RGB = ColorTranslator.ToOle(Color.FromArgb(255, 255, 255));
            shape.Line.Transparency = 0.35f;
            shape.Glow.Radius = 0f;
            shape.Glow.Transparency = 1f;
            shape.SoftEdge.Radius = 0f;
        }

        private static void RemoveLegacyHelperOutlines(PowerPoint.Shape shape)
        {
            PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
            if (slide == null)
            {
                return;
            }

            for (int index = slide.Shapes.Count; index >= 1; index--)
            {
                PowerPoint.Shape candidate = slide.Shapes[index];
                if (candidate.Tags[GlassHelperTag] == "1")
                {
                    candidate.Delete();
                }
            }
        }

        private static void ApplyShadow(PowerPoint.Shape shape)
        {
            shape.Shadow.Visible = MsoTriState.msoTrue;
            shape.Shadow.OffsetX = 1.5f;
            shape.Shadow.OffsetY = 1.5f;
            shape.Shadow.Transparency = 0.80f;
            shape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(191, 191, 191));
            shape.Shadow.Size = 1.01f;
            shape.Shadow.Blur = 0f;
        }

        private static void ApplyBevel(PowerPoint.Shape shape)
        {
            shape.ThreeD.Visible = MsoTriState.msoTrue;
            shape.ThreeD.BevelTopType = MsoBevelType.msoBevelCircle;
            shape.ThreeD.BevelTopInset = 5f;
            shape.ThreeD.BevelTopDepth = 1f;
            shape.ThreeD.ContourWidth = 0f;
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
