using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class LiquidGlassEffectService
    {
        private static readonly Color DarkShadowColor = Color.FromArgb(105, 105, 105);
        private static readonly Color LightShadowColor = Color.FromArgb(210, 210, 210);
        private readonly PowerPoint.Application application;

        public LiquidGlassEffectService(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public int ApplyToSelection(bool debugMode)
        {
            return ApplyToSelection(debugMode, null);
        }

        public int ApplyLayeredToSelection(bool debugMode, Color overlayColor)
        {
            return ApplyToSelection(debugMode, overlayColor);
        }

        private int ApplyToSelection(bool debugMode, Color? overlayColor)
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
                changedCount += ApplyToShapeRecursive(shapeRange[index], debugMode, overlayColor);
            }

            if (changedCount == 0)
            {
                throw new InvalidOperationException("No supported shapes were found in the current selection.");
            }

            return changedCount;
        }

        private int ApplyToShapeRecursive(PowerPoint.Shape shape, bool debugMode, Color? overlayColor)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                return 0;
            }

            if (!SupportsLiquidGlass(shape))
            {
                return 0;
            }

            DebugStep(debugMode, $"Applying liquid-glass effect to: {shape.Name}");

            if (overlayColor.HasValue)
            {
                ApplyLayeredGlass(shape, overlayColor.Value);
            }
            else
            {
                ApplyStandardGlass(shape);
            }

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

        private static void ApplyStandardGlass(PowerPoint.Shape shape)
        {
            ApplyBackgroundFill(shape);
            ClearOutline(shape);
            ClearEffects(shape);
            ApplyBevel(shape);
            ApplyShadow(shape, DarkShadowColor);
        }

        private static void ApplyLayeredGlass(PowerPoint.Shape shape, Color overlayColor)
        {
            ApplyBackgroundFill(shape);
            ClearOutline(shape);
            ClearEffects(shape);
            shape.Shadow.Visible = MsoTriState.msoFalse;

            PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
            if (slide == null)
            {
                throw new InvalidOperationException("The selected shape must be on a slide.");
            }

            string groupKey = Guid.NewGuid().ToString("N");
            shape.Name = $"PPTAssistant Base {groupKey}";

            PowerPoint.Shape overlay = shape.Duplicate()[1];
            overlay.Name = $"PPTAssistant Overlay {groupKey}";
            overlay.Fill.Visible = MsoTriState.msoTrue;
            overlay.Fill.Solid();
            overlay.Fill.ForeColor.RGB = ColorTranslator.ToOle(overlayColor);
            overlay.Fill.Transparency = 0.75f;
            overlay.Line.Visible = MsoTriState.msoFalse;
            overlay.Glow.Radius = 0f;
            overlay.Glow.Transparency = 1f;
            overlay.SoftEdge.Radius = 0f;
            overlay.ThreeD.Visible = MsoTriState.msoFalse;
            overlay.Left = shape.Left;
            overlay.Top = shape.Top;
            overlay.Width = shape.Width;
            overlay.Height = shape.Height;
            ApplyShadow(overlay, overlayColor == Color.Black ? LightShadowColor : DarkShadowColor);

            PowerPoint.Shape grouped = slide.Shapes.Range(new object[] { shape.Name, overlay.Name }).Group();
            grouped.Name = $"PPTAssistant Glass {groupKey}";
        }

        private static void ApplyBackgroundFill(PowerPoint.Shape shape)
        {
            shape.Fill.Visible = MsoTriState.msoTrue;
            shape.Fill.Background();
        }

        private static void ClearOutline(PowerPoint.Shape shape)
        {
            shape.Line.Visible = MsoTriState.msoFalse;
        }

        private static void ClearEffects(PowerPoint.Shape shape)
        {
            shape.Glow.Radius = 0f;
            shape.Glow.Transparency = 1f;
            shape.SoftEdge.Radius = 0f;
        }

        private static void ApplyShadow(PowerPoint.Shape shape, Color shadowColor)
        {
            shape.Shadow.Visible = MsoTriState.msoTrue;
            shape.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            shape.Shadow.Type = MsoShadowType.msoShadow5;
            shape.Shadow.OffsetX = 0f;
            shape.Shadow.OffsetY = 0f;
            shape.Shadow.Transparency = 0.80f;
            shape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(shadowColor);
            shape.Shadow.Size = 1.01f;
            shape.Shadow.Blur = 0f;
        }

        private static void ApplyBevel(PowerPoint.Shape shape)
        {
            shape.ThreeD.Visible = MsoTriState.msoTrue;
            shape.ThreeD.BevelTopType = MsoBevelType.msoBevelCircle;
            shape.ThreeD.BevelTopInset = 20f;
            shape.ThreeD.BevelTopDepth = 1f;
            shape.ThreeD.ContourWidth = 0f;
            shape.ThreeD.PresetMaterial = MsoPresetMaterial.msoMaterialSoftEdge;
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
