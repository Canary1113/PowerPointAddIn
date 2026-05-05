using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class ThemeTemplateInfo
    {
        public ThemeTemplateInfo(string label, int slideIndex, string tagValue)
        {
            Label = label;
            SlideIndex = slideIndex;
            TagValue = tagValue;
        }

        public string Label { get; }

        public int SlideIndex { get; }

        public string TagValue { get; }
    }

    internal sealed class ThemeSectionInfo
    {
        public ThemeSectionInfo(string name, IReadOnlyList<int> slideIndexes)
        {
            Name = name;
            SlideIndexes = slideIndexes;
        }

        public string Name { get; }

        public IReadOnlyList<int> SlideIndexes { get; }
    }

    internal sealed class ThemePackageInfo
    {
        public ThemePackageInfo(int slideCount, IReadOnlyList<ThemeSectionInfo> sections, IReadOnlyDictionary<int, string> slideVariants)
        {
            SlideCount = slideCount;
            Sections = sections;
            SlideVariants = slideVariants;
        }

        public int SlideCount { get; }

        public IReadOnlyList<ThemeSectionInfo> Sections { get; }

        public IReadOnlyDictionary<int, string> SlideVariants { get; }
    }

    internal sealed class ThemeService
    {
        private const string ThemeShapeTag = "PPTAssistantTheme";
        private const string ThemeTemplateFileName = "ThemeStyles.pptx";
        private static readonly Color DefaultAccentColor = Color.FromArgb(0, 112, 192);
        private static Color currentAccentColor = DefaultAccentColor;
        private const float PointsPerCentimeter = 28.3464567f;
        private const float SlideWidthCentimeters = 33.8667f;
        private const float Style2HeaderHeightCentimeters = 2.33f;
        private const float Style2HeaderWidthCentimeters = 33.93f;
        private const float NavWidthCentimeters = 15f;
        private const float NavHeightCentimeters = 1f;
        private const float Style1NavTopCentimeters = 0.7838f;
        private const float Style2NavTopCentimeters = 0.6147f;
        private const float SelectedWidthCentimeters = 2.5f;
        private const float SelectedHeightCentimeters = 0.8f;
        private const float Style1SelectedLeftInsetCentimeters = 0.1201f;
        private const float Style2SelectedLeftInsetCentimeters = 0.155f;
        private const float SelectedTopInsetCentimeters = 0.1f;
        private readonly PowerPoint.Application application;

        public ThemeService(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public IReadOnlyList<ThemeTemplateInfo> GetAvailableTemplates()
        {
            string templatePath = ResolveThemeTemplatePath();
            using (ZipArchive archive = ZipFile.OpenRead(templatePath))
            {
                ThemePackageInfo packageInfo = ReadThemePackageInfo(archive);
                List<ThemeTemplateInfo> templates = ReadSectionTemplates(packageInfo);
                if (templates.Count == 0)
                {
                    templates = ReadSlideTemplates(packageInfo);
                }

                return templates;
            }
        }

        public void ApplyTemplate(ThemeTemplateInfo template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            ApplyTemplateSlide(template.SlideIndex, template.TagValue);
        }

        public void ApplyStyle1()
        {
            ApplyTemplateSlide(1, "Style1Light");
        }

        public void ApplyStyle2()
        {
            ApplyTemplateSlide(3, "Style2Light");
        }

        public void ApplyStyle1Light()
        {
            ApplyTemplateSlide(1, "Style1Light");
        }

        public void ApplyStyle1Dark()
        {
            ApplyTemplateSlide(2, "Style1Dark");
        }

        public void ApplyStyle2Light()
        {
            ApplyTemplateSlide(3, "Style2Light");
        }

        public void ApplyStyle2Dark()
        {
            ApplyTemplateSlide(4, "Style2Dark");
        }

        private void ApplyTemplateSlide(int slideIndex, string tagValue)
        {
            PowerPointShapeContext context = PowerPointShapeContext.FromActiveView(application);
            if (!PrepareSurfaceForTheme(context))
            {
                return;
            }

            PasteTemplateStyle(context, slideIndex, tagValue);
        }

        private static ThemePackageInfo ReadThemePackageInfo(ZipArchive archive)
        {
            ZipArchiveEntry presentationEntry = archive.GetEntry("ppt/presentation.xml");
            if (presentationEntry == null)
            {
                throw new InvalidOperationException("The theme template is missing ppt/presentation.xml.");
            }

            XDocument presentationDocument;
            using (Stream stream = presentationEntry.Open())
            {
                presentationDocument = XDocument.Load(stream);
            }

            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
            var slideIdToIndex = new Dictionary<uint, int>();
            var relationshipIdToIndex = new Dictionary<string, int>();
            int slideIndex = 1;

            foreach (XElement slideId in presentationDocument.Descendants(p + "sldId"))
            {
                XAttribute idAttribute = slideId.Attribute("id");
                XAttribute relationshipAttribute = slideId.Attribute(r + "id");
                if (idAttribute != null && uint.TryParse(idAttribute.Value, out uint slideIdValue))
                {
                    slideIdToIndex[slideIdValue] = slideIndex;
                }

                if (relationshipAttribute != null)
                {
                    relationshipIdToIndex[relationshipAttribute.Value] = slideIndex;
                }

                slideIndex++;
            }

            var slideVariants = new Dictionary<int, string>();
            for (int index = 1; index < slideIndex; index++)
            {
                slideVariants[index] = DetectSlideVariant(archive, index);
            }

            var sections = new List<ThemeSectionInfo>();
            foreach (XElement section in presentationDocument.Descendants(p14 + "section"))
            {
                XAttribute nameAttribute = section.Attribute("name");
                string sectionName = string.IsNullOrWhiteSpace(nameAttribute?.Value) ? $"Section {sections.Count + 1}" : nameAttribute.Value;
                var slideIndexes = new List<int>();

                foreach (XElement sectionSlideId in section.Descendants(p14 + "sldId"))
                {
                    XAttribute idAttribute = sectionSlideId.Attribute("id");
                    if (idAttribute != null &&
                        uint.TryParse(idAttribute.Value, out uint sectionSlideIdValue) &&
                        slideIdToIndex.TryGetValue(sectionSlideIdValue, out int sectionSlideIndex))
                    {
                        slideIndexes.Add(sectionSlideIndex);
                    }
                }

                if (slideIndexes.Count > 0)
                {
                    sections.Add(new ThemeSectionInfo(sectionName, slideIndexes));
                }
            }

            return new ThemePackageInfo(slideIndex - 1, sections, slideVariants);
        }

        private static List<ThemeTemplateInfo> ReadSectionTemplates(ThemePackageInfo packageInfo)
        {
            var templates = new List<ThemeTemplateInfo>();

            foreach (ThemeSectionInfo section in packageInfo.Sections)
            {
                for (int index = 0; index < section.SlideIndexes.Count; index++)
                {
                    int slideIndex = section.SlideIndexes[index];
                    string label = section.Name;
                    if (section.SlideIndexes.Count > 1)
                    {
                        packageInfo.SlideVariants.TryGetValue(slideIndex, out string variant);
                        label = string.IsNullOrEmpty(variant)
                            ? $"{section.Name} {index + 1}"
                            : $"{section.Name} {variant}";
                    }

                    templates.Add(new ThemeTemplateInfo(label, slideIndex, ToTagValue(label, templates.Count + 1)));
                }
            }

            return templates;
        }

        private static string DetectSlideVariant(ZipArchive archive, int slideIndex)
        {
            ZipArchiveEntry slideEntry = archive.GetEntry($"ppt/slides/slide{slideIndex}.xml");
            if (slideEntry == null)
            {
                return string.Empty;
            }

            XDocument slideDocument;
            using (Stream stream = slideEntry.Open())
            {
                slideDocument = XDocument.Load(stream);
            }

            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XElement firstShape = slideDocument.Descendants(p + "sp").FirstOrDefault();
            XElement fill = firstShape?.Descendants(a + "solidFill").FirstOrDefault();
            XElement srgb = fill?.Element(a + "srgbClr");
            XAttribute colorAttribute = srgb?.Attribute("val");
            if (colorAttribute != null && TryParseHexColor(colorAttribute.Value, out Color color))
            {
                return GetVariantName(color);
            }

            XElement scheme = fill?.Element(a + "schemeClr");
            string schemeValue = scheme?.Attribute("val")?.Value;
            string lumModValue = scheme?.Element(a + "lumMod")?.Attribute("val")?.Value;
            if (schemeValue == "tx1")
            {
                return "Dark";
            }

            if (schemeValue == "bg1" || schemeValue == "bg2")
            {
                if (int.TryParse(lumModValue, out int lumMod) && lumMod < 60000)
                {
                    return "Dark";
                }

                return "Light";
            }

            return string.Empty;
        }

        private static bool TryParseHexColor(string value, out Color color)
        {
            color = Color.Empty;
            if (value == null || value.Length != 6)
            {
                return false;
            }

            try
            {
                int red = Convert.ToInt32(value.Substring(0, 2), 16);
                int green = Convert.ToInt32(value.Substring(2, 2), 16);
                int blue = Convert.ToInt32(value.Substring(4, 2), 16);
                color = Color.FromArgb(red, green, blue);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static string GetVariantName(Color color)
        {
            double luminance = (0.2126 * color.R) + (0.7152 * color.G) + (0.0722 * color.B);
            return luminance < 128d ? "Dark" : "Light";
        }

        private static List<ThemeTemplateInfo> ReadSlideTemplates(ThemePackageInfo packageInfo)
        {
            var templates = new List<ThemeTemplateInfo>();
            for (int slideIndex = 1; slideIndex <= packageInfo.SlideCount; slideIndex++)
            {
                templates.Add(new ThemeTemplateInfo($"Slide {slideIndex}", slideIndex, $"Slide{slideIndex}"));
            }

            return templates;
        }

        private static string ToTagValue(string label, int index)
        {
            string compactLabel = string.Empty;
            foreach (char character in label)
            {
                if (char.IsLetterOrDigit(character))
                {
                    compactLabel += character;
                }
            }

            if (string.IsNullOrEmpty(compactLabel))
            {
                compactLabel = "Section";
            }

            return compactLabel + index.ToString();
        }

        public void CustomizeAccentColor(Color accentColor)
        {
            PowerPointShapeContext context = PowerPointShapeContext.FromActiveView(application);
            int defaultAccentOle = ColorTranslator.ToOle(DefaultAccentColor);
            int currentAccentOle = ColorTranslator.ToOle(currentAccentColor);
            int replacementOle = ColorTranslator.ToOle(accentColor);

            for (int index = 1; index <= context.Shapes.Count; index++)
            {
                PowerPoint.Shape shape = context.Shapes[index];
                if (IsThemeShape(shape))
                {
                    ApplyAccentColor(shape, defaultAccentOle, currentAccentOle, replacementOle);
                }
            }

            currentAccentColor = accentColor;
        }

        public void OpenThemeTemplate()
        {
            string templatePath = ResolveThemeTemplatePath();
            application.Presentations.Open(
                templatePath,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue);
        }

        private static bool PrepareSurfaceForTheme(PowerPointShapeContext context)
        {
            if (context.Shapes.Count == 0)
            {
                return true;
            }

            ExistingContentAction action = ShowExistingContentDialog();
            if (action == ExistingContentAction.Delete)
            {
                context.ClearShapes();
            }

            return true;
        }

        private void PasteTemplateStyle(PowerPointShapeContext context, int slideIndex, string tagValue)
        {
            string templatePath = ResolveThemeTemplatePath();
            PowerPoint.Presentation templatePresentation = null;
            PowerPoint.Presentation targetPresentation = application.ActivePresentation;
            float targetSlideWidth = targetPresentation.PageSetup.SlideWidth;
            float targetSlideHeight = targetPresentation.PageSetup.SlideHeight;

            try
            {
                templatePresentation = application.Presentations.Open(
                    templatePath,
                    MsoTriState.msoTrue,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);

                if (templatePresentation.Slides.Count < slideIndex)
                {
                    throw new InvalidOperationException($"The theme template does not contain slide {slideIndex}.");
                }

                PowerPoint.Slide sourceSlide = templatePresentation.Slides[slideIndex];
                if (sourceSlide.Shapes.Count == 0)
                {
                    return;
                }

                IReadOnlyList<PowerPoint.Shape> pastedShapes;
                try
                {
                    pastedShapes = CopyTemplateShapes(sourceSlide, context.Shapes);
                }
                catch (COMException)
                {
                    templatePresentation.Close();
                    templatePresentation = null;
                    pastedShapes = CopyTemplateShapesViaImportedSlide(
                        targetPresentation,
                        context.Shapes,
                        templatePath,
                        slideIndex);
                }

                TagPastedShapes(pastedShapes, tagValue);
                SendFullSlideBackgroundToBack(pastedShapes, targetSlideWidth, targetSlideHeight);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("PowerPoint could not copy the requested theme template.", ex);
            }
            finally
            {
                if (templatePresentation != null)
                {
                    templatePresentation.Close();
                }
            }
        }

        private static IReadOnlyList<PowerPoint.Shape> CopyTemplateShapes(PowerPoint.Slide sourceSlide, PowerPoint.Shapes targetShapes)
        {
            var pastedShapes = new List<PowerPoint.Shape>();
            try
            {
                dynamic sourceShapes = sourceSlide.Shapes;
                int[] shapeIndexes = Enumerable.Range(1, sourceSlide.Shapes.Count).ToArray();
                PowerPoint.ShapeRange sourceRange = sourceShapes.Range(shapeIndexes);
                sourceRange.Copy();
                AddPastedShapes(targetShapes.Paste(), pastedShapes);
            }
            catch (COMException)
            {
                CopyTemplateShapesOneByOne(sourceSlide, targetShapes, pastedShapes);
            }

            if (pastedShapes.Count == 0)
            {
                throw new COMException("No template shapes could be copied.");
            }

            return pastedShapes;
        }

        private static IReadOnlyList<PowerPoint.Shape> CopyTemplateShapesViaImportedSlide(
            PowerPoint.Presentation targetPresentation,
            PowerPoint.Shapes targetShapes,
            string templatePath,
            int slideIndex)
        {
            var pastedShapes = new List<PowerPoint.Shape>();
            PowerPoint.Slide importedSlide = null;

            try
            {
                int insertAfter = targetPresentation.Slides.Count;
                targetPresentation.Slides.InsertFromFile(templatePath, insertAfter, slideIndex, slideIndex);
                importedSlide = targetPresentation.Slides[insertAfter + 1];
                CopyTemplateShapesOneByOne(importedSlide, targetShapes, pastedShapes);

                if (pastedShapes.Count == 0)
                {
                    throw new COMException("No imported template shapes could be copied.");
                }

                return pastedShapes;
            }
            finally
            {
                if (importedSlide != null)
                {
                    try
                    {
                        importedSlide.Delete();
                    }
                    catch (COMException)
                    {
                    }
                }
            }
        }

        private static void CopyTemplateShapesOneByOne(
            PowerPoint.Slide sourceSlide,
            PowerPoint.Shapes targetShapes,
            ICollection<PowerPoint.Shape> pastedShapes)
        {
            for (int index = 1; index <= sourceSlide.Shapes.Count; index++)
            {
                try
                {
                    sourceSlide.Shapes[index].Copy();
                    AddPastedShapes(targetShapes.Paste(), pastedShapes);
                }
                catch (COMException)
                {
                }
            }
        }

        private static void AddPastedShapes(PowerPoint.ShapeRange pastedRange, ICollection<PowerPoint.Shape> pastedShapes)
        {
            for (int index = 1; index <= pastedRange.Count; index++)
            {
                pastedShapes.Add(pastedRange[index]);
            }
        }

        private static void ApplyAccentColor(PowerPoint.Shape shape, int defaultAccentOle, int currentAccentOle, int replacementOle)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                for (int index = 1; index <= shape.GroupItems.Count; index++)
                {
                    ApplyAccentColor(shape.GroupItems[index], defaultAccentOle, currentAccentOle, replacementOle);
                }

                return;
            }

            TryReplaceLineColor(shape, defaultAccentOle, currentAccentOle, replacementOle, IsAccentLineShape(shape));
            TryReplaceTextColor(shape, defaultAccentOle, currentAccentOle, replacementOle);
        }

        private static bool IsThemeShape(PowerPoint.Shape shape)
        {
            try
            {
                return !string.IsNullOrEmpty(shape.Tags[ThemeShapeTag]);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static void TryReplaceLineColor(PowerPoint.Shape shape, int defaultAccentOle, int currentAccentOle, int replacementOle, bool forceReplace)
        {
            try
            {
                if (shape.Line.Visible == MsoTriState.msoTrue && (forceReplace || IsAccentColor(shape.Line.ForeColor.RGB, defaultAccentOle, currentAccentOle)))
                {
                    shape.Line.ForeColor.RGB = replacementOle;
                }
            }
            catch (COMException)
            {
            }
        }

        private static void TryReplaceTextColor(PowerPoint.Shape shape, int defaultAccentOle, int currentAccentOle, int replacementOle)
        {
            try
            {
                if (shape.HasTextFrame != MsoTriState.msoTrue || shape.TextFrame.HasText != MsoTriState.msoTrue)
                {
                    return;
                }

                PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                if (TryReplaceKnownAccentText(shape, textRange, replacementOle))
                {
                    return;
                }

                for (int index = 1; index <= textRange.Characters().Count; index++)
                {
                    PowerPoint.TextRange character = textRange.Characters(index, 1);
                    if (IsAccentColor(character.Font.Color.RGB, defaultAccentOle, currentAccentOle))
                    {
                        character.Font.Color.RGB = replacementOle;
                    }
                }
            }
            catch (COMException)
            {
            }
        }

        private static bool TryReplaceKnownAccentText(PowerPoint.Shape shape, PowerPoint.TextRange textRange, int replacementOle)
        {
            string shapeName = shape.Name ?? string.Empty;
            string text = textRange.Text ?? string.Empty;

            if (shapeName.IndexOf("Slide Number", StringComparison.OrdinalIgnoreCase) >= 0 ||
                shapeName.IndexOf("灯片编号", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                textRange.Font.Color.RGB = replacementOle;
                return true;
            }

            int overviewIndex = text.IndexOf("Overview", StringComparison.OrdinalIgnoreCase);
            if (overviewIndex >= 0)
            {
                textRange.Characters(overviewIndex + 1, "Overview".Length).Font.Color.RGB = replacementOle;
                return false;
            }

            return false;
        }

        private static bool IsAccentLineShape(PowerPoint.Shape shape)
        {
            string shapeName = shape.Name ?? string.Empty;
            if (shapeName.IndexOf("PageArc", StringComparison.OrdinalIgnoreCase) >= 0 ||
                shapeName.IndexOf("弧形", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            try
            {
                return shape.AutoShapeType == MsoAutoShapeType.msoShapeArc;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static bool IsAccentColor(int colorOle, int defaultAccentOle, int currentAccentOle)
        {
            return colorOle == defaultAccentOle || colorOle == currentAccentOle;
        }

        private static ExistingContentAction ShowExistingContentDialog()
        {
            const int deleteButtonId = 1001;
            const int keepButtonId = 1002;

            var buttons = new[]
            {
                new TaskDialogButton { ButtonId = deleteButtonId, ButtonText = "Delete" },
                new TaskDialogButton { ButtonId = keepButtonId, ButtonText = "Keep" }
            };

            IntPtr buttonsPointer = IntPtr.Zero;
            try
            {
                buttonsPointer = AllocateTaskDialogButtons(buttons);
                var config = new TaskDialogConfig
                {
                    Size = (uint)Marshal.SizeOf(typeof(TaskDialogConfig)),
                    WindowTitle = "Apply Theme",
                    MainInstruction = "This page already contains content.",
                    Content = "Delete existing content before applying the theme, or keep it and add the theme to the current page.",
                    ButtonCount = (uint)buttons.Length,
                    Buttons = buttonsPointer,
                    DefaultButton = keepButtonId
                };

                TaskDialogIndirect(ref config, out int selectedButton, out _, out _);
                return selectedButton == deleteButtonId ? ExistingContentAction.Delete : ExistingContentAction.Keep;
            }
            catch (Exception)
            {
                DialogResult fallbackResult = MessageBox.Show(
                    "This page already contains content.\r\n\r\nDelete existing content?\r\nChoose No to keep existing content.",
                    "Apply Theme",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);

                return fallbackResult == DialogResult.Yes ? ExistingContentAction.Delete : ExistingContentAction.Keep;
            }
            finally
            {
                if (buttonsPointer != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(buttonsPointer);
                }
            }
        }

        private enum ExistingContentAction
        {
            Keep,
            Delete
        }

        private static IntPtr AllocateTaskDialogButtons(TaskDialogButton[] buttons)
        {
            int buttonSize = Marshal.SizeOf(typeof(TaskDialogButton));
            IntPtr buttonsPointer = Marshal.AllocHGlobal(buttonSize * buttons.Length);

            for (int index = 0; index < buttons.Length; index++)
            {
                IntPtr itemPointer = IntPtr.Add(buttonsPointer, index * buttonSize);
                Marshal.StructureToPtr(buttons[index], itemPointer, false);
            }

            return buttonsPointer;
        }

        [DllImport("comctl32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void TaskDialogIndirect(
            ref TaskDialogConfig config,
            out int button,
            out int radioButton,
            [MarshalAs(UnmanagedType.Bool)] out bool verificationFlagChecked);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct TaskDialogButton
        {
            public int ButtonId;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string ButtonText;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct TaskDialogConfig
        {
            public uint Size;
            public IntPtr ParentWindow;
            public IntPtr Instance;
            public uint Flags;
            public uint CommonButtons;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string WindowTitle;

            public IntPtr MainIcon;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string MainInstruction;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string Content;

            public uint ButtonCount;
            public IntPtr Buttons;
            public int DefaultButton;
            public uint RadioButtonCount;
            public IntPtr RadioButtons;
            public int DefaultRadioButton;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string VerificationText;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string ExpandedInformation;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string ExpandedControlText;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string CollapsedControlText;

            public IntPtr FooterIcon;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string Footer;

            public IntPtr Callback;
            public IntPtr CallbackData;
            public uint Width;
        }

        private static void TagPastedShapes(IEnumerable<PowerPoint.Shape> pastedShapes, string tagValue)
        {
            foreach (PowerPoint.Shape shape in pastedShapes)
            {
                try
                {
                    shape.Tags.Add(ThemeShapeTag, tagValue);
                }
                catch (COMException)
                {
                }
            }
        }

        private static void SendFullSlideBackgroundToBack(IEnumerable<PowerPoint.Shape> pastedShapes, float slideWidth, float slideHeight)
        {
            foreach (PowerPoint.Shape shape in pastedShapes)
            {
                bool isFullSlide =
                    Math.Abs(shape.Left) < 1f &&
                    Math.Abs(shape.Top) < 1f &&
                    Math.Abs(shape.Width - slideWidth) < 2f &&
                    Math.Abs(shape.Height - slideHeight) < 2f;

                if (isFullSlide)
                {
                    shape.ZOrder(MsoZOrderCmd.msoSendToBack);
                    return;
                }
            }
        }

        private static string ResolveThemeTemplatePath()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDirectory, "Assets", ThemeTemplateFileName),
                Path.GetFullPath(Path.Combine(baseDirectory, "..", "..", "Assets", ThemeTemplateFileName)),
                Path.GetFullPath(Path.Combine(baseDirectory, "..", "..", "..", "PowerPointAddIn", "Assets", ThemeTemplateFileName))
            };

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            throw new FileNotFoundException("The theme template file was not found.", candidates[0]);
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
            PowerPointShapeContext.TrySetShapeLocked(background, MsoTriState.msoTrue);
            background.ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        private static void CreateStyle1Navigation(PowerPointShapeContext context)
        {
            float navLeft = CenteredLeft(NavWidthCentimeters);
            PowerPoint.Shape capsule = AddRoundedRectangle(context, "Style1Capsule", navLeft, Style1NavTopCentimeters, NavWidthCentimeters, NavHeightCentimeters);
            capsule.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            capsule.Fill.Transparency = 0.30f;
            capsule.Line.Visible = MsoTriState.msoTrue;
            capsule.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            capsule.Line.Weight = 0.75f;
            ApplyCenteredShadow(capsule, Color.FromArgb(191, 191, 191), 0.60f, 102f, 5f);
            PowerPointShapeContext.TrySetShapeLocked(capsule, MsoTriState.msoFalse);

            PowerPoint.Shape selected = AddRoundedRectangle(
                context,
                "Style1SelectedCapsule",
                navLeft + Style1SelectedLeftInsetCentimeters,
                Style1NavTopCentimeters + SelectedTopInsetCentimeters,
                SelectedWidthCentimeters,
                SelectedHeightCentimeters);
            selected.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(209, 209, 209));
            selected.Fill.Transparency = 0.70f;
            selected.Line.Visible = MsoTriState.msoFalse;
            selected.Shadow.Visible = MsoTriState.msoFalse;
            PowerPointShapeContext.TrySetShapeLocked(selected, MsoTriState.msoFalse);

            AddNavText(context, navLeft, Style1NavTopCentimeters, navLeft + Style1SelectedLeftInsetCentimeters);
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
            header.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
            header.Line.Weight = 2f;
            header.Shadow.Visible = MsoTriState.msoTrue;
            header.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            header.Shadow.Type = MsoShadowType.msoShadow21;
            header.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
            header.Shadow.Transparency = 0.50f;
            header.Shadow.Blur = 4f;
            header.Shadow.OffsetX = 0f;
            header.Shadow.OffsetY = 3f;
            PowerPointShapeContext.TrySetShapeLocked(header, MsoTriState.msoTrue);
        }

        private static void CreateStyle2Navigation(PowerPointShapeContext context)
        {
            float navLeft = CenteredLeft(NavWidthCentimeters);
            PowerPoint.Shape capsule = AddRoundedRectangle(context, "Style2Capsule", navLeft, Style2NavTopCentimeters, NavWidthCentimeters, NavHeightCentimeters);
            capsule.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(234, 234, 234));
            capsule.Fill.Transparency = 0f;
            capsule.Line.Visible = MsoTriState.msoFalse;
            capsule.Shadow.Visible = MsoTriState.msoFalse;
            PowerPointShapeContext.TrySetShapeLocked(capsule, MsoTriState.msoFalse);

            PowerPoint.Shape selected = AddRoundedRectangle(
                context,
                "Style2SelectedCapsule",
                navLeft + Style2SelectedLeftInsetCentimeters,
                Style2NavTopCentimeters + SelectedTopInsetCentimeters,
                SelectedWidthCentimeters,
                SelectedHeightCentimeters);
            selected.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            selected.Fill.Transparency = 0f;
            selected.Line.Visible = MsoTriState.msoFalse;
            ApplyCenteredShadow(selected, Color.FromArgb(196, 196, 196), 0.60f, 102f, 5f);
            PowerPointShapeContext.TrySetShapeLocked(selected, MsoTriState.msoFalse);

            AddNavText(context, navLeft, Style2NavTopCentimeters, navLeft + Style2SelectedLeftInsetCentimeters);
        }

        private static void CreateStyle1Title(PowerPointShapeContext context)
        {
            PowerPoint.Shape titleCapsule = AddRoundedRectangle(context, "Style1TitleCapsule", 2.0123f, 0.782f, 2.9977f, 1f);
            titleCapsule.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            titleCapsule.Fill.Transparency = 0.30f;
            titleCapsule.Line.Visible = MsoTriState.msoTrue;
            titleCapsule.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            titleCapsule.Line.Weight = 0.75f;
            ApplyCenteredShadow(titleCapsule, Color.FromArgb(191, 191, 191), 0.40f, 101f, 5f);
            SetShapeText(titleCapsule, "Project Brief", Color.FromArgb(89, 89, 89));
            PowerPointShapeContext.TrySetShapeLocked(titleCapsule, MsoTriState.msoFalse);
        }

        private static void CreateStyle2Title(PowerPointShapeContext context)
        {
            AddPlainText(context, "Style2Title", "Project Brief", 0.2208f, 0.716f, 3.2742f, 0.7694f, Color.FromArgb(89, 89, 89), PowerPoint.PpParagraphAlignment.ppAlignCenter);
        }

        private static void CreateStyle1PageMark(PowerPointShapeContext context)
        {
            PowerPoint.Shape circle = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                CentimetersToPoints(31.8527f),
                CentimetersToPoints(0.8372f),
                CentimetersToPoints(1.0031f),
                CentimetersToPoints(1.0031f));
            circle.Name = $"PPTAssistant Theme Style1 Page Circle {DateTime.Now:HHmmss}";
            circle.Tags.Add(ThemeShapeTag, "Style1PageCircle");
            circle.Fill.Visible = MsoTriState.msoTrue;
            circle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            circle.Fill.Transparency = 0.30f;
            circle.Line.Visible = MsoTriState.msoTrue;
            circle.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            circle.Line.Weight = 0.75f;
            ApplyCenteredShadow(circle, Color.FromArgb(128, 128, 128), 0.60f, 102f, 5f);
            AddSlideNumberToShape(circle);
            PowerPointShapeContext.TrySetShapeLocked(circle, MsoTriState.msoFalse);

            AddPageArc(context, "Style1PageArc", 31.97f, 0.9525f, 0.7686f, 0.80f);
        }

        private static void CreateStyle2PageMark(PowerPointShapeContext context)
        {
            AddSlideNumberText(context, "Style2SlideNumber", 32.0443f, 0.6437f, 1.0809f, 1.0142f);
            AddPageArc(context, "Style2PageArc", 32.2005f, 0.7666f, 0.7686f, 0.60f);
        }

        private static PowerPoint.Shape AddRoundedRectangle(PowerPointShapeContext context, string role, float leftCm, float topCm, float widthCm, float heightCm)
        {
            PowerPoint.Shape shape = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                CentimetersToPoints(leftCm),
                CentimetersToPoints(topCm),
                CentimetersToPoints(widthCm),
                CentimetersToPoints(heightCm));

            shape.Name = $"PPTAssistant Theme {role} {DateTime.Now:HHmmss}";
            shape.Tags.Add(ThemeShapeTag, role);
            shape.Fill.Visible = MsoTriState.msoTrue;
            shape.Line.Visible = MsoTriState.msoFalse;
            shape.Adjustments[1] = 1f;
            return shape;
        }

        private static void AddNavText(PowerPointShapeContext context, float navLeftCm, float navTopCm, float selectedLeftCm)
        {
            AddPlainText(context, "NavTextOverview", "Overview", selectedLeftCm, navTopCm, SelectedWidthCentimeters, NavHeightCentimeters, Color.FromArgb(0, 112, 192), PowerPoint.PpParagraphAlignment.ppAlignCenter);
            AddPlainText(context, "NavTextBackground", "Background", navLeftCm + 3.45f, navTopCm, 2.4f, NavHeightCentimeters, Color.FromArgb(64, 64, 64), PowerPoint.PpParagraphAlignment.ppAlignCenter);
            AddPlainText(context, "NavTextImplementation", "Implementation", navLeftCm + 6.20f, navTopCm, 4.2f, NavHeightCentimeters, Color.FromArgb(64, 64, 64), PowerPoint.PpParagraphAlignment.ppAlignCenter);
            AddPlainText(context, "NavTextAnalysis", "Analysis", navLeftCm + 10.35f, navTopCm, 2.0f, NavHeightCentimeters, Color.FromArgb(64, 64, 64), PowerPoint.PpParagraphAlignment.ppAlignCenter);
            AddPlainText(context, "NavTextConclusion", "Conclusion", navLeftCm + 12.35f, navTopCm, 2.65f, NavHeightCentimeters, Color.FromArgb(64, 64, 64), PowerPoint.PpParagraphAlignment.ppAlignCenter);
        }

        private static PowerPoint.Shape AddPlainText(PowerPointShapeContext context, string role, string text, float leftCm, float topCm, float widthCm, float heightCm, Color color, PowerPoint.PpParagraphAlignment alignment)
        {
            PowerPoint.Shape shape = context.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                CentimetersToPoints(leftCm),
                CentimetersToPoints(topCm),
                CentimetersToPoints(widthCm),
                CentimetersToPoints(heightCm));

            shape.Name = $"PPTAssistant Theme {role} {DateTime.Now:HHmmss}";
            shape.Tags.Add(ThemeShapeTag, role);
            shape.Fill.Visible = MsoTriState.msoFalse;
            shape.Line.Visible = MsoTriState.msoFalse;
            PrepareTextFrame(shape, alignment);
            shape.TextFrame.TextRange.Text = text;
            FormatTextRange(shape.TextFrame.TextRange, color);
            PowerPointShapeContext.TrySetShapeLocked(shape, MsoTriState.msoFalse);
            return shape;
        }

        private static void AddSlideNumberText(PowerPointShapeContext context, string role, float leftCm, float topCm, float widthCm, float heightCm)
        {
            PowerPoint.Shape number = AddPlainText(context, role, string.Empty, leftCm, topCm, widthCm, heightCm, Color.FromArgb(0, 112, 192), PowerPoint.PpParagraphAlignment.ppAlignCenter);
            number.TextFrame.TextRange.InsertSlideNumber();
            number.TextFrame.TextRange.Font.Name = "Calibri";
            number.TextFrame.TextRange.Font.Size = 12f;
            number.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            number.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
        }

        private static void AddSlideNumberToShape(PowerPoint.Shape shape)
        {
            PrepareTextFrame(shape, PowerPoint.PpParagraphAlignment.ppAlignCenter);
            shape.TextFrame.TextRange.Text = string.Empty;
            shape.TextFrame.TextRange.InsertSlideNumber();
            FormatTextRange(shape.TextFrame.TextRange, Color.FromArgb(0, 112, 192));
        }

        private static void AddPageArc(PowerPointShapeContext context, string role, float leftCm, float topCm, float sizeCm, float transparency)
        {
            PowerPoint.Shape arc = context.Shapes.AddShape(
                MsoAutoShapeType.msoShapeArc,
                CentimetersToPoints(leftCm),
                CentimetersToPoints(topCm),
                CentimetersToPoints(sizeCm),
                CentimetersToPoints(sizeCm));

            arc.Name = $"PPTAssistant Theme {role} {DateTime.Now:HHmmss}";
            arc.Tags.Add(ThemeShapeTag, role);
            arc.Fill.Visible = MsoTriState.msoFalse;
            arc.Line.Visible = MsoTriState.msoTrue;
            arc.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
            arc.Line.Transparency = transparency;
            arc.Line.Weight = 2.5f;
            TrySetReferenceArcAngles(arc);
            PowerPointShapeContext.TrySetShapeLocked(arc, MsoTriState.msoFalse);
        }

        private static void SetShapeText(PowerPoint.Shape shape, string text, Color color)
        {
            PrepareTextFrame(shape, PowerPoint.PpParagraphAlignment.ppAlignCenter);
            shape.TextFrame.TextRange.Text = text;
            FormatTextRange(shape.TextFrame.TextRange, color);
        }

        private static void PrepareTextFrame(PowerPoint.Shape shape, PowerPoint.PpParagraphAlignment alignment)
        {
            shape.TextFrame.MarginLeft = 0f;
            shape.TextFrame.MarginRight = 0f;
            shape.TextFrame.MarginTop = 0f;
            shape.TextFrame.MarginBottom = 0f;
            shape.TextFrame.WordWrap = MsoTriState.msoFalse;
            shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = alignment;
        }

        private static void FormatTextRange(PowerPoint.TextRange range, Color color)
        {
            range.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            range.Font.Name = "Calibri";
            range.Font.Size = 12f;
            range.Font.Bold = MsoTriState.msoTrue;
            range.Font.Color.RGB = ColorTranslator.ToOle(color);
        }

        private static void TrySetReferenceArcAngles(PowerPoint.Shape arc)
        {
            try
            {
                arc.Adjustments[1] = 7336547f / 60000f;
                arc.Adjustments[2] = 3412435f / 60000f;
            }
            catch (COMException)
            {
            }
        }

        private static void ApplyCenteredShadow(PowerPoint.Shape shape, Color color, float transparency, float size, float blur)
        {
            shape.Shadow.Visible = MsoTriState.msoTrue;
            shape.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            shape.Shadow.Type = MsoShadowType.msoShadow21;
            shape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(color);
            shape.Shadow.Transparency = transparency;
            shape.Shadow.Size = size;
            shape.Shadow.Blur = blur;
            shape.Shadow.OffsetX = 0f;
            shape.Shadow.OffsetY = 0f;
        }

        private static float CenteredLeft(float widthCentimeters)
        {
            return (SlideWidthCentimeters - widthCentimeters) / 2f;
        }

        private static float CentimetersToPoints(float centimeters)
        {
            return centimeters * PointsPerCentimeter;
        }
    }
}
