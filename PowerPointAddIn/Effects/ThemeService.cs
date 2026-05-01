using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class ThemeService
    {
        private const string ThemeShapeTag = "PPTAssistantTheme";
        private const float PointsPerCentimeter = 28.3464567f;
        private const float Style2HeaderHeightCentimeters = 2.33f;
        private const float Style2HeaderWidthCentimeters = 33.93f;
        private const float Style2CapsuleWidthCentimeters = 16.3f;
        private const float Style2CapsuleHeightCentimeters = 1f;
        private const float Style1SelectedCapsuleWidthCentimeters = 3.67f;
        private const float Style1SelectedCapsuleHeightCentimeters = 1.21f;
        private const float Style2SelectedCapsuleWidthCentimeters = 3.22f;
        private const float Style2SelectedCapsuleHeightCentimeters = 0.72f;
        private const float Style2PageNumberTopCentimeters = 0.73f;
        private const float Style2PageNumberRightCentimeters = 1.05f;
        private readonly PowerPoint.Application application;

        public ThemeService(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void ApplyStyle1()
        {
            PowerPointShapeContext context = PowerPointShapeContext.FromActiveView(application);
            context.ClearShapesIfTemplateSurface();
            CreateGrayBackground(context, "Style1");
            CreateStyle1Capsule(context);
            CreateStyle1SelectedCapsule(context);
        }

        public void ApplyStyle2()
        {
            PowerPointShapeContext context = PowerPointShapeContext.FromActiveView(application);
            context.ClearShapesIfTemplateSurface();
            CreateGrayBackground(context, "Style2Background");
            CreateStyle2Header(context);
            CreateStyle2Capsule(context);
            CreateStyle2SelectedCapsule(context);
            CreateStyle2PageNumber(context);
        }

        private void CreateGrayBackground(PowerPointShapeContext context, string tagValue)
        {
            float slideWidth = application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = application.ActivePresentation.PageSetup.SlideHeight;

            PowerPoint.Shape background = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                0f,
                0f,
                slideWidth,
                slideHeight);

            background.Name = $"PPTAssistant Theme {tagValue} {DateTime.Now:HHmmss}";
            background.Tags.Add(ThemeShapeTag, tagValue);
            background.Fill.Visible = MsoTriState.msoTrue;
            background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(247, 247, 247));
            background.Line.Visible = MsoTriState.msoFalse;
            background.Locked = MsoTriState.msoTrue;
            background.ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        private static void CreateStyle2Header(PowerPointShapeContext context)
        {
            PowerPoint.Shape header = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                0f,
                0f,
                CentimetersToPoints(Style2HeaderWidthCentimeters),
                CentimetersToPoints(Style2HeaderHeightCentimeters));

            header.Name = $"PPTAssistant Theme Style2 Header {DateTime.Now:HHmmss}";
            header.Tags.Add(ThemeShapeTag, "Style2Header");
            header.Fill.Visible = MsoTriState.msoTrue;
            header.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            header.Line.Visible = MsoTriState.msoTrue;
            header.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            header.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
            header.Line.Weight = 2f;
            header.Shadow.Visible = MsoTriState.msoTrue;
            header.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            header.Shadow.Type = MsoShadowType.msoShadow21;
            header.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            header.Shadow.Transparency = 0.86f;
            header.Shadow.Blur = 8f;
            header.Shadow.OffsetX = 0f;
            header.Shadow.OffsetY = 2f;
            header.Locked = MsoTriState.msoTrue;
        }

        private static void CreateStyle2Capsule(PowerPointShapeContext context)
        {
            float headerWidth = CentimetersToPoints(Style2HeaderWidthCentimeters);
            float headerHeight = CentimetersToPoints(Style2HeaderHeightCentimeters);
            float capsuleWidth = CentimetersToPoints(Style2CapsuleWidthCentimeters);
            float capsuleHeight = CentimetersToPoints(Style2CapsuleHeightCentimeters);

            PowerPoint.Shape capsule = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                (headerWidth - capsuleWidth) / 2f,
                (headerHeight - capsuleHeight) / 2f,
                capsuleWidth,
                capsuleHeight);

            capsule.Name = $"PPTAssistant Theme Style2 Capsule {DateTime.Now:HHmmss}";
            capsule.Tags.Add(ThemeShapeTag, "Style2Capsule");
            capsule.Fill.Visible = MsoTriState.msoTrue;
            capsule.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(234, 234, 234));
            capsule.Fill.Transparency = 0f;
            capsule.Line.Visible = MsoTriState.msoFalse;
            capsule.Shadow.Visible = MsoTriState.msoFalse;
            capsule.Adjustments[1] = 1f;
            capsule.Locked = MsoTriState.msoTrue;
        }

        private static void CreateStyle1Capsule(PowerPointShapeContext context)
        {
            float headerWidth = CentimetersToPoints(Style2HeaderWidthCentimeters);
            float headerHeight = CentimetersToPoints(Style2HeaderHeightCentimeters);
            float capsuleWidth = CentimetersToPoints(Style2CapsuleWidthCentimeters);
            float capsuleHeight = CentimetersToPoints(Style2CapsuleHeightCentimeters);

            PowerPoint.Shape capsule = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                (headerWidth - capsuleWidth) / 2f,
                (headerHeight - capsuleHeight) / 2f,
                capsuleWidth,
                capsuleHeight);

            capsule.Name = $"PPTAssistant Theme Style1 Capsule {DateTime.Now:HHmmss}";
            capsule.Tags.Add(ThemeShapeTag, "Style1Capsule");
            capsule.Fill.Visible = MsoTriState.msoTrue;
            capsule.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            capsule.Fill.Transparency = 0.30f;
            capsule.Line.Visible = MsoTriState.msoTrue;
            capsule.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            capsule.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            capsule.Line.Weight = 0.75f;
            capsule.Shadow.Visible = MsoTriState.msoTrue;
            capsule.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            capsule.Shadow.Type = MsoShadowType.msoShadow21;
            capsule.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(191, 191, 191));
            capsule.Shadow.Transparency = 0.60f;
            capsule.Shadow.Size = 101f;
            capsule.Shadow.Blur = 5f;
            capsule.Shadow.OffsetX = 0f;
            capsule.Shadow.OffsetY = 0f;
            capsule.Adjustments[1] = 1f;
            capsule.Locked = MsoTriState.msoFalse;
        }

        private static void CreateStyle1SelectedCapsule(PowerPointShapeContext context)
        {
            float headerWidth = CentimetersToPoints(Style2HeaderWidthCentimeters);
            float headerHeight = CentimetersToPoints(Style2HeaderHeightCentimeters);
            float capsuleWidth = CentimetersToPoints(Style2CapsuleWidthCentimeters);
            float capsuleHeight = CentimetersToPoints(Style2CapsuleHeightCentimeters);
            float selectedWidth = CentimetersToPoints(Style1SelectedCapsuleWidthCentimeters);
            float selectedHeight = CentimetersToPoints(Style1SelectedCapsuleHeightCentimeters);
            float capsuleLeft = (headerWidth - capsuleWidth) / 2f;
            float capsuleTop = (headerHeight - capsuleHeight) / 2f;

            PowerPoint.Shape selectedCapsule = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                capsuleLeft,
                capsuleTop + (capsuleHeight - selectedHeight) / 2f,
                selectedWidth,
                selectedHeight);

            selectedCapsule.Name = $"PPTAssistant Theme Style1 Selected Capsule {DateTime.Now:HHmmss}";
            selectedCapsule.Tags.Add(ThemeShapeTag, "Style1SelectedCapsule");
            selectedCapsule.Fill.Visible = MsoTriState.msoFalse;
            selectedCapsule.Line.Visible = MsoTriState.msoTrue;
            selectedCapsule.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            selectedCapsule.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            selectedCapsule.Line.Transparency = 0.75f;
            selectedCapsule.Line.Weight = 0.5f;
            selectedCapsule.Shadow.Visible = MsoTriState.msoTrue;
            selectedCapsule.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            selectedCapsule.Shadow.Type = MsoShadowType.msoShadow21;
            selectedCapsule.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            selectedCapsule.Shadow.Transparency = 0.60f;
            selectedCapsule.Shadow.Size = 101f;
            selectedCapsule.Shadow.Blur = 5f;
            selectedCapsule.Shadow.OffsetX = 0f;
            selectedCapsule.Shadow.OffsetY = 0f;
            selectedCapsule.Adjustments[1] = 1f;
            selectedCapsule.Locked = MsoTriState.msoFalse;
        }

        private static void CreateStyle2SelectedCapsule(PowerPointShapeContext context)
        {
            float headerWidth = CentimetersToPoints(Style2HeaderWidthCentimeters);
            float headerHeight = CentimetersToPoints(Style2HeaderHeightCentimeters);
            float capsuleWidth = CentimetersToPoints(Style2CapsuleWidthCentimeters);
            float capsuleHeight = CentimetersToPoints(Style2CapsuleHeightCentimeters);
            float selectedWidth = CentimetersToPoints(Style2SelectedCapsuleWidthCentimeters);
            float selectedHeight = CentimetersToPoints(Style2SelectedCapsuleHeightCentimeters);
            float capsuleLeft = (headerWidth - capsuleWidth) / 2f;
            float capsuleTop = (headerHeight - capsuleHeight) / 2f;

            PowerPoint.Shape selectedCapsule = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                capsuleLeft,
                capsuleTop + (capsuleHeight - selectedHeight) / 2f,
                selectedWidth,
                selectedHeight);

            selectedCapsule.Name = $"PPTAssistant Theme Style2 Selected Capsule {DateTime.Now:HHmmss}";
            selectedCapsule.Tags.Add(ThemeShapeTag, "Style2SelectedCapsule");
            selectedCapsule.Fill.Visible = MsoTriState.msoTrue;
            selectedCapsule.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            selectedCapsule.Fill.Transparency = 0f;
            selectedCapsule.Line.Visible = MsoTriState.msoFalse;
            selectedCapsule.Shadow.Visible = MsoTriState.msoTrue;
            selectedCapsule.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            selectedCapsule.Shadow.Type = MsoShadowType.msoShadow21;
            selectedCapsule.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(196, 196, 196));
            selectedCapsule.Shadow.Transparency = 0.60f;
            selectedCapsule.Shadow.Size = 102f;
            selectedCapsule.Shadow.Blur = 5f;
            selectedCapsule.Shadow.OffsetX = 0f;
            selectedCapsule.Shadow.OffsetY = 0f;
            selectedCapsule.Adjustments[1] = 1f;
            selectedCapsule.Locked = MsoTriState.msoFalse;
        }

        private static void CreateStyle2PageNumber(PowerPointShapeContext context)
        {
            float headerWidth = CentimetersToPoints(Style2HeaderWidthCentimeters);
            float top = CentimetersToPoints(Style2PageNumberTopCentimeters);
            float right = CentimetersToPoints(Style2PageNumberRightCentimeters);
            float height = CentimetersToPoints(0.52f);
            float suffixWidth = CentimetersToPoints(0.32f);
            float numberWidth = CentimetersToPoints(0.42f);
            float prefixWidth = CentimetersToPoints(0.42f);
            float suffixLeft = headerWidth - right - suffixWidth;
            float numberLeft = suffixLeft - numberWidth;
            float prefixLeft = numberLeft - prefixWidth;

            PowerPoint.Shape prefix = AddPageText(context, "Style2PagePrefix", "\u7b2c", prefixLeft, top, prefixWidth, height, Color.FromArgb(156, 156, 156));
            PowerPoint.Shape number = AddPageText(context, "Style2PageNumber", string.Empty, numberLeft, top, numberWidth, height, Color.FromArgb(229, 45, 38));
            PowerPoint.Shape suffix = AddPageText(context, "Style2PageSuffix", "\u9875", suffixLeft, top, suffixWidth, height, Color.FromArgb(156, 156, 156));

            number.TextFrame.TextRange.InsertSlideNumber();
            FormatPageText(number, Color.FromArgb(229, 45, 38));
            prefix.Locked = MsoTriState.msoTrue;
            number.Locked = MsoTriState.msoTrue;
            suffix.Locked = MsoTriState.msoTrue;
        }

        private static PowerPoint.Shape AddPageText(
            PowerPointShapeContext context,
            string role,
            string text,
            float left,
            float top,
            float width,
            float height,
            Color color)
        {
            PowerPoint.Shape shape = context.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                width,
                height);

            shape.Name = $"PPTAssistant Theme {role} {DateTime.Now:HHmmss}";
            shape.Tags.Add(ThemeShapeTag, role);
            shape.Fill.Visible = MsoTriState.msoFalse;
            shape.Line.Visible = MsoTriState.msoFalse;
            shape.TextFrame.MarginLeft = 0f;
            shape.TextFrame.MarginRight = 0f;
            shape.TextFrame.MarginTop = 0f;
            shape.TextFrame.MarginBottom = 0f;
            shape.TextFrame.WordWrap = MsoTriState.msoFalse;
            shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            shape.TextFrame.TextRange.Text = text;
            FormatPageText(shape, color);

            return shape;
        }

        private static void FormatPageText(PowerPoint.Shape shape, Color color)
        {
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            shape.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
            shape.TextFrame.TextRange.Font.Size = 12f;
            shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(color);
        }

        private static float CentimetersToPoints(float centimeters)
        {
            return centimeters * PointsPerCentimeter;
        }
    }
}
