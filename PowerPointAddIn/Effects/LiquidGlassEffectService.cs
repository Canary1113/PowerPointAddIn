using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class LiquidGlassEffectService
    {
        private const string BorderTagName = "PPTAssistantBorder";
        private static readonly Color ShadowColor = Color.FromArgb(90, 90, 90);
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
            ApplyShadow(shape);
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
            ApplyShadow(overlay);

            PowerPoint.Shape border = CreateGradientBorder(slide, overlay, groupKey);
            border.ZOrder(MsoZOrderCmd.msoBringToFront);

            PowerPoint.Shape grouped = slide.Shapes.Range(new object[] { shape.Name, overlay.Name, border.Name }).Group();
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

        private static void ApplyShadow(PowerPoint.Shape shape)
        {
            shape.Shadow.Visible = MsoTriState.msoTrue;
            shape.Shadow.OffsetX = 1.5f;
            shape.Shadow.OffsetY = 1.5f;
            shape.Shadow.Transparency = 0.80f;
            shape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(ShadowColor);
            shape.Shadow.Size = 1.01f;
            shape.Shadow.Blur = 0f;
        }

        private static PowerPoint.Shape CreateGradientBorder(PowerPoint.Slide slide, PowerPoint.Shape referenceShape, string groupKey)
        {
            const float lineWeight = 0.5f;
            float halfWeight = lineWeight / 2f;

            PowerPoint.Shape outer = referenceShape.Duplicate()[1];
            PowerPoint.Shape inner = referenceShape.Duplicate()[1];
            string borderKey = $"PPTAssistant Border {groupKey}";

            outer.Name = $"{borderKey} Outer";
            outer.Tags.Add(BorderTagName, borderKey);
            outer.Left -= halfWeight;
            outer.Top -= halfWeight;
            outer.Width += lineWeight;
            outer.Height += lineWeight;
            outer.Line.Visible = MsoTriState.msoFalse;
            outer.Shadow.Visible = MsoTriState.msoFalse;
            outer.ThreeD.Visible = MsoTriState.msoFalse;
            outer.Glow.Radius = 0f;
            outer.SoftEdge.Radius = 0f;

            inner.Name = $"{borderKey} Inner";
            inner.Left += halfWeight;
            inner.Top += halfWeight;
            inner.Width = Math.Max(1f, inner.Width - lineWeight);
            inner.Height = Math.Max(1f, inner.Height - lineWeight);
            inner.Line.Visible = MsoTriState.msoFalse;
            inner.Fill.Visible = MsoTriState.msoTrue;
            inner.Fill.Solid();
            inner.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            inner.Fill.Transparency = 0f;
            inner.Shadow.Visible = MsoTriState.msoFalse;
            inner.ThreeD.Visible = MsoTriState.msoFalse;
            inner.Glow.Radius = 0f;
            inner.SoftEdge.Radius = 0f;

            outer.Fill.Visible = MsoTriState.msoTrue;
            outer.Fill.OneColorGradient(MsoGradientStyle.msoGradientDiagonalDown, 1, 1f);
            outer.Fill.GradientAngle = 45f;
            ResetGradientStops(outer.Fill);
            AddGradientStop(outer.Fill, 0.00f, 0.10f);
            AddGradientStop(outer.Fill, 0.40f, 0.40f);
            AddGradientStop(outer.Fill, 0.70f, 0.10f);
            AddGradientStop(outer.Fill, 1.00f, 0.90f);

            slide.Shapes.Range(new object[] { outer.Name, inner.Name }).MergeShapes(MsoMergeCmd.msoMergeSubtract, outer);

            PowerPoint.Shape border = FindTaggedShape(slide, borderKey)
                ?? throw new InvalidOperationException("Failed to create the gradient border.");
            border.Name = $"{borderKey} Ring";
            border.Line.Visible = MsoTriState.msoFalse;
            border.Shadow.Visible = MsoTriState.msoFalse;
            border.ThreeD.Visible = MsoTriState.msoFalse;
            border.Glow.Radius = 0f;
            border.SoftEdge.Radius = 0f;
            return border;
        }

        private static void ResetGradientStops(PowerPoint.FillFormat fill)
        {
            for (int index = fill.GradientStops.Count; index >= 1; index--)
            {
                fill.GradientStops.Delete(index);
            }
        }

        private static void AddGradientStop(PowerPoint.FillFormat fill, float position, float transparency)
        {
            fill.GradientStops.Insert(
                ColorTranslator.ToOle(Color.White),
                position,
                transparency,
                0);
        }

        private static PowerPoint.Shape FindTaggedShape(PowerPoint.Slide slide, string borderKey)
        {
            for (int index = 1; index <= slide.Shapes.Count; index++)
            {
                PowerPoint.Shape candidate = slide.Shapes[index];
                if (candidate.Tags[BorderTagName] == borderKey)
                {
                    return candidate;
                }
            }

            return null;
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
