using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class LiquidGlassEffectService
    {
        private const string GlassTag = "PPTAssistantGlass";
        private const string GlassKindTag = "PPTAssistantGlassKind";
        private const string GlassRoleTag = "PPTAssistantGlassRole";
        private const string GlassGroupKeyTag = "PPTAssistantGlassGroupKey";
        private const string OriginalNameTag = "PPTAssistantOriginalName";
        private const string OriginalFillVisibleTag = "PPTAssistantOriginalFillVisible";
        private const string OriginalFillColorTag = "PPTAssistantOriginalFillColor";
        private const string OriginalFillTransparencyTag = "PPTAssistantOriginalFillTransparency";
        private const string OriginalLineVisibleTag = "PPTAssistantOriginalLineVisible";
        private const string OriginalLineColorTag = "PPTAssistantOriginalLineColor";
        private const string OriginalLineTransparencyTag = "PPTAssistantOriginalLineTransparency";
        private const string OriginalLineWeightTag = "PPTAssistantOriginalLineWeight";
        private const string StandardGlassKind = "Standard";
        private const string LayeredGlassKind = "Layered";
        private const string SingleRole = "Single";
        private const string BaseRole = "Base";
        private const string OverlayRole = "Overlay";
        private const string GroupRole = "Group";
        private const float ShadowScalePercent = 101f;
        private static readonly Color UnifiedShadowColor = Color.FromArgb(115, 115, 115);
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

        public int RecoverSelection(bool debugMode)
        {
            PowerPoint.ShapeRange shapeRange = GetSelectedShapeRange();
            int changedCount = 0;

            for (int index = 1; index <= shapeRange.Count; index++)
            {
                changedCount += RecoverShape(shapeRange[index], debugMode);
            }

            if (changedCount == 0)
            {
                throw new InvalidOperationException("No recoverable Glass shapes were found in the current selection.");
            }

            return changedCount;
        }

        private int ApplyToSelection(bool debugMode, Color? overlayColor)
        {
            PowerPoint.ShapeRange shapeRange = GetSelectedShapeRange();
            int changedCount = 0;

            for (int index = 1; index <= shapeRange.Count; index++)
            {
                changedCount += ApplyToShape(shapeRange[index], debugMode, overlayColor);
            }

            if (changedCount == 0)
            {
                throw new InvalidOperationException("No supported shapes were found in the current selection.");
            }

            return changedCount;
        }

        private PowerPoint.ShapeRange GetSelectedShapeRange()
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

            return selection.ShapeRange;
        }

        private int ApplyToShape(PowerPoint.Shape shape, bool debugMode, Color? overlayColor)
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

        private int RecoverShape(PowerPoint.Shape shape, bool debugMode)
        {
            if (IsLayeredGlassGroup(shape))
            {
                DebugStep(debugMode, $"Recovering layered glass from: {shape.Name}");
                RecoverLayeredGlassGroup(shape);
                return 1;
            }

            if (shape.Type == MsoShapeType.msoGroup)
            {
                return 0;
            }

            if (!IsRecoverableStandardGlass(shape))
            {
                return 0;
            }

            DebugStep(debugMode, $"Recovering glass effect from: {shape.Name}");
            RecoverStandardGlass(shape);
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
            CaptureOriginalAppearance(shape);
            ApplyBackgroundFill(shape);
            ClearOutline(shape);
            ClearEffects(shape);
            ApplyBevel(shape);
            ApplyShadow(shape, UnifiedShadowColor);
            MarkGlass(shape, StandardGlassKind, SingleRole, groupKey: null);
        }

        private static void ApplyLayeredGlass(PowerPoint.Shape shape, Color overlayColor)
        {
            CaptureOriginalAppearance(shape);
            ApplyBackgroundFill(shape);
            ClearOutline(shape);
            ClearEffects(shape);
            ApplyShadow(shape, UnifiedShadowColor);

            PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
            if (slide == null)
            {
                throw new InvalidOperationException("The selected shape must be on a slide.");
            }

            string groupKey = Guid.NewGuid().ToString("N");
            shape.Name = $"PPTAssistant Base {groupKey}";
            MarkGlass(shape, LayeredGlassKind, BaseRole, groupKey);

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
            overlay.Shadow.Visible = MsoTriState.msoFalse;
            ClearGlassState(overlay);
            ClearOriginalAppearanceTags(overlay);
            MarkGlass(overlay, LayeredGlassKind, OverlayRole, groupKey);

            PowerPoint.Shape grouped = slide.Shapes.Range(new object[] { shape.Name, overlay.Name }).Group();
            grouped.Name = $"PPTAssistant Glass {groupKey}";
            ApplyBevel(grouped);
            grouped.Shadow.Visible = MsoTriState.msoFalse;
            MarkGlass(grouped, LayeredGlassKind, GroupRole, groupKey);
        }

        private static bool IsRecoverableStandardGlass(PowerPoint.Shape shape)
        {
            return IsTaggedGlassShape(shape, StandardGlassKind, SingleRole) || LooksLikeLegacyStandardGlass(shape);
        }

        private static bool IsLayeredGlassGroup(PowerPoint.Shape shape)
        {
            if (shape.Type != MsoShapeType.msoGroup)
            {
                return false;
            }

            return IsTaggedGlassShape(shape, LayeredGlassKind, GroupRole) ||
                   shape.Name.StartsWith("PPTAssistant Glass ", StringComparison.Ordinal);
        }

        private static bool IsLayeredGlassBase(PowerPoint.Shape shape)
        {
            return IsTaggedGlassShape(shape, LayeredGlassKind, BaseRole) ||
                   shape.Name.StartsWith("PPTAssistant Base ", StringComparison.Ordinal);
        }

        private static bool IsLayeredGlassOverlay(PowerPoint.Shape shape)
        {
            return IsTaggedGlassShape(shape, LayeredGlassKind, OverlayRole) ||
                   shape.Name.StartsWith("PPTAssistant Overlay ", StringComparison.Ordinal);
        }

        private static bool LooksLikeLegacyStandardGlass(PowerPoint.Shape shape)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                return false;
            }

            return shape.Shadow.Visible == MsoTriState.msoTrue &&
                   shape.ThreeD.Visible == MsoTriState.msoTrue &&
                   shape.ThreeD.BevelTopType == MsoBevelType.msoBevelCircle &&
                   Math.Abs(shape.ThreeD.BevelTopInset - 10f) < 0.01f &&
                   Math.Abs(shape.ThreeD.BevelTopDepth - 1f) < 0.01f &&
                   shape.Line.Visible == MsoTriState.msoFalse;
        }

        private static void RecoverLayeredGlassGroup(PowerPoint.Shape groupedShape)
        {
            PowerPoint.ShapeRange members = groupedShape.Ungroup();
            PowerPoint.Shape baseShape = null;
            PowerPoint.Shape overlayShape = null;

            for (int index = 1; index <= members.Count; index++)
            {
                PowerPoint.Shape member = members[index];
                if (baseShape == null && IsLayeredGlassBase(member))
                {
                    baseShape = member;
                    continue;
                }

                if (overlayShape == null && IsLayeredGlassOverlay(member))
                {
                    overlayShape = member;
                }
            }

            if (baseShape == null || overlayShape == null)
            {
                throw new InvalidOperationException("The selected Glass group could not be recovered safely.");
            }

            overlayShape.Delete();
            RecoverStandardGlass(baseShape);
        }

        private static void RecoverStandardGlass(PowerPoint.Shape shape)
        {
            ClearGlassState(shape);
            RestoreFill(shape);
            RestoreLine(shape);
            RestoreOriginalName(shape);
            ClearGlassTags(shape);
            ClearOriginalAppearanceTags(shape);
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
            shape.Shadow.Type = MsoShadowType.msoShadow25;
            shape.Shadow.Size = ShadowScalePercent;
            shape.Shadow.Transparency = 0.80f;
            shape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(shadowColor);
        }

        private static void ApplyBevel(PowerPoint.Shape shape)
        {
            shape.ThreeD.Visible = MsoTriState.msoTrue;
            shape.ThreeD.BevelTopType = MsoBevelType.msoBevelCircle;
            shape.ThreeD.BevelTopInset = 10f;
            shape.ThreeD.BevelTopDepth = 1f;
            shape.ThreeD.ContourWidth = 0f;
            shape.ThreeD.PresetMaterial = MsoPresetMaterial.msoMaterialSoftEdge;
        }

        private static void ClearGlassState(PowerPoint.Shape shape)
        {
            shape.Shadow.Visible = MsoTriState.msoFalse;
            shape.ThreeD.Visible = MsoTriState.msoFalse;
            shape.ThreeD.BevelTopType = MsoBevelType.msoBevelNone;
            shape.ThreeD.BevelTopInset = 0f;
            shape.ThreeD.BevelTopDepth = 0f;
            shape.ThreeD.ContourWidth = 0f;
            ClearEffects(shape);
        }

        private static void CaptureOriginalAppearance(PowerPoint.Shape shape)
        {
            if (!string.IsNullOrEmpty(shape.Tags[OriginalNameTag]))
            {
                return;
            }

            SetTag(shape, OriginalNameTag, shape.Name);
            SetTag(shape, OriginalFillVisibleTag, ToTagBool(shape.Fill.Visible == MsoTriState.msoTrue));
            SetTag(shape, OriginalLineVisibleTag, ToTagBool(shape.Line.Visible == MsoTriState.msoTrue));

            if (shape.Fill.Visible == MsoTriState.msoTrue)
            {
                SetTag(shape, OriginalFillColorTag, shape.Fill.ForeColor.RGB.ToString(CultureInfo.InvariantCulture));
                SetTag(shape, OriginalFillTransparencyTag, shape.Fill.Transparency.ToString(CultureInfo.InvariantCulture));
            }

            if (shape.Line.Visible == MsoTriState.msoTrue)
            {
                SetTag(shape, OriginalLineColorTag, shape.Line.ForeColor.RGB.ToString(CultureInfo.InvariantCulture));
                SetTag(shape, OriginalLineTransparencyTag, shape.Line.Transparency.ToString(CultureInfo.InvariantCulture));
                SetTag(shape, OriginalLineWeightTag, shape.Line.Weight.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static void RestoreFill(PowerPoint.Shape shape)
        {
            if (!TryGetTagBool(shape, OriginalFillVisibleTag, out bool fillVisible))
            {
                return;
            }

            shape.Fill.Visible = fillVisible ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            if (!fillVisible)
            {
                return;
            }

            if (TryGetTagInt(shape, OriginalFillColorTag, out int fillColor))
            {
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = fillColor;
            }

            if (TryGetTagFloat(shape, OriginalFillTransparencyTag, out float fillTransparency))
            {
                shape.Fill.Transparency = fillTransparency;
            }
        }

        private static void RestoreLine(PowerPoint.Shape shape)
        {
            if (!TryGetTagBool(shape, OriginalLineVisibleTag, out bool lineVisible))
            {
                return;
            }

            shape.Line.Visible = lineVisible ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            if (!lineVisible)
            {
                return;
            }

            if (TryGetTagInt(shape, OriginalLineColorTag, out int lineColor))
            {
                shape.Line.ForeColor.RGB = lineColor;
            }

            if (TryGetTagFloat(shape, OriginalLineTransparencyTag, out float lineTransparency))
            {
                shape.Line.Transparency = lineTransparency;
            }

            if (TryGetTagFloat(shape, OriginalLineWeightTag, out float lineWeight))
            {
                shape.Line.Weight = lineWeight;
            }
        }

        private static void RestoreOriginalName(PowerPoint.Shape shape)
        {
            string originalName = shape.Tags[OriginalNameTag];
            if (string.IsNullOrEmpty(originalName))
            {
                return;
            }

            try
            {
                shape.Name = originalName;
            }
            catch
            {
            }
        }

        private static void MarkGlass(PowerPoint.Shape shape, string kind, string role, string groupKey)
        {
            SetTag(shape, GlassTag, "1");
            SetTag(shape, GlassKindTag, kind);
            SetTag(shape, GlassRoleTag, role);

            if (string.IsNullOrEmpty(groupKey))
            {
                DeleteTag(shape, GlassGroupKeyTag);
            }
            else
            {
                SetTag(shape, GlassGroupKeyTag, groupKey);
            }
        }

        private static bool IsTaggedGlassShape(PowerPoint.Shape shape, string kind, string role)
        {
            return shape.Tags[GlassTag] == "1" &&
                   shape.Tags[GlassKindTag] == kind &&
                   shape.Tags[GlassRoleTag] == role;
        }

        private static void ClearGlassTags(PowerPoint.Shape shape)
        {
            DeleteTag(shape, GlassTag);
            DeleteTag(shape, GlassKindTag);
            DeleteTag(shape, GlassRoleTag);
            DeleteTag(shape, GlassGroupKeyTag);
        }

        private static void ClearOriginalAppearanceTags(PowerPoint.Shape shape)
        {
            DeleteTag(shape, OriginalNameTag);
            DeleteTag(shape, OriginalFillVisibleTag);
            DeleteTag(shape, OriginalFillColorTag);
            DeleteTag(shape, OriginalFillTransparencyTag);
            DeleteTag(shape, OriginalLineVisibleTag);
            DeleteTag(shape, OriginalLineColorTag);
            DeleteTag(shape, OriginalLineTransparencyTag);
            DeleteTag(shape, OriginalLineWeightTag);
        }

        private static void SetTag(PowerPoint.Shape shape, string name, string value)
        {
            DeleteTag(shape, name);
            shape.Tags.Add(name, value);
        }

        private static void DeleteTag(PowerPoint.Shape shape, string name)
        {
            try
            {
                shape.Tags.Delete(name);
            }
            catch
            {
            }
        }

        private static string ToTagBool(bool value)
        {
            return value ? "1" : "0";
        }

        private static bool TryGetTagBool(PowerPoint.Shape shape, string tagName, out bool value)
        {
            value = shape.Tags[tagName] == "1";
            return shape.Tags[tagName] == "1" || shape.Tags[tagName] == "0";
        }

        private static bool TryGetTagInt(PowerPoint.Shape shape, string tagName, out int value)
        {
            return int.TryParse(shape.Tags[tagName], NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryGetTagFloat(PowerPoint.Shape shape, string tagName, out float value)
        {
            return float.TryParse(shape.Tags[tagName], NumberStyles.Float, CultureInfo.InvariantCulture, out value);
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
