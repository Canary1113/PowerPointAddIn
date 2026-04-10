using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class NavbarAnimationService
    {
        private const float MoveDurationSeconds = 1f;
        private const float GrowDurationSeconds = 0.5f;
        private const float ShrinkDurationSeconds = 0.5f;
        private const float ShrinkDelaySeconds = 0.5f;
        private const float GrowScalePercent = 125f;
        private const float ShrinkScalePercent = 80f;
        private readonly PowerPoint.Application application;

        public NavbarAnimationService(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public int ApplyToSelection()
        {
            PowerPoint.ShapeRange shapeRange = GetSelectedShapeRange();
            PowerPoint.Presentation presentation = application.ActivePresentation;
            EnsurePresentationSupportsAbsoluteScalePatch(presentation);
            int changedCount = 0;
            var targetShapeIds = new List<int>();
            PowerPoint.Slide targetSlide = null;

            for (int index = 1; index <= shapeRange.Count; index++)
            {
                PowerPoint.Shape shape = shapeRange[index];
                PowerPoint.Slide shapeSlide = shape.Parent as PowerPoint.Slide;
                if (shapeSlide == null)
                {
                    throw new InvalidOperationException("The selected shape must be on a slide.");
                }

                if (targetSlide == null)
                {
                    targetSlide = shapeSlide;
                }
                else if (shapeSlide.SlideIndex != targetSlide.SlideIndex)
                {
                    throw new InvalidOperationException("Select shapes from only one slide at a time.");
                }

                ApplyToShape(shape);
                targetShapeIds.Add(shape.Id);
                changedCount++;
            }

            if (targetSlide != null)
            {
                ApplyAbsoluteScalePatch(presentation, targetSlide.SlideIndex, targetShapeIds);
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
                throw new InvalidOperationException("Select one or more navigation shapes first.");
            }

            return selection.ShapeRange;
        }

        private void ApplyToShape(PowerPoint.Shape shape)
        {
            PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
            if (slide == null)
            {
                throw new InvalidOperationException("The selected shape must be on a slide.");
            }

            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            RemoveExistingEffects(shape, sequence);

            float slideWidth = application.ActivePresentation.PageSetup.SlideWidth;
            if (slideWidth <= 0f)
            {
                throw new InvalidOperationException("The active presentation has an invalid slide width.");
            }

            float moveDistanceRatio = shape.Width / slideWidth;

            PowerPoint.Effect moveEffect = sequence.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectPathRight);
            moveEffect.Timing.Duration = MoveDurationSeconds;
            moveEffect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            moveEffect.Timing.Accelerate = 0f;
            moveEffect.Timing.Decelerate = 0f;
            moveEffect.Behaviors[1].MotionEffect.Path = CreateHorizontalLinePath(moveDistanceRatio);

            PowerPoint.Effect growEffect = sequence.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink);
            growEffect.Timing.Duration = GrowDurationSeconds;
            growEffect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            growEffect.Timing.Accelerate = 0f;
            growEffect.Timing.Decelerate = 0f;
            growEffect.Behaviors[1].ScaleEffect.ToX = GrowScalePercent;
            growEffect.Behaviors[1].ScaleEffect.ToY = GrowScalePercent;

            PowerPoint.Effect shrinkEffect = sequence.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink);
            shrinkEffect.Timing.Duration = ShrinkDurationSeconds;
            shrinkEffect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            shrinkEffect.Timing.TriggerDelayTime = ShrinkDelaySeconds;
            shrinkEffect.Timing.Accelerate = 0f;
            shrinkEffect.Timing.Decelerate = 0f;
            shrinkEffect.Behaviors[1].ScaleEffect.ToX = ShrinkScalePercent;
            shrinkEffect.Behaviors[1].ScaleEffect.ToY = ShrinkScalePercent;
        }

        private static string CreateHorizontalLinePath(float moveDistanceRatio)
        {
            float clampedRatio = Math.Max(0.001f, moveDistanceRatio);
            return string.Format(
                CultureInfo.InvariantCulture,
                "M 0 0 L {0:0.######} 0 E",
                clampedRatio);
        }

        private static void RemoveExistingEffects(PowerPoint.Shape shape, PowerPoint.Sequence sequence)
        {
            for (int index = sequence.Count; index >= 1; index--)
            {
                PowerPoint.Effect effect = sequence[index];
                if (effect.Shape != null && effect.Shape.Id == shape.Id)
                {
                    effect.Delete();
                }
            }
        }

        private static void EnsurePresentationSupportsAbsoluteScalePatch(PowerPoint.Presentation presentation)
        {
            if (presentation == null || string.IsNullOrEmpty(presentation.FullName))
            {
                throw new InvalidOperationException("Save the presentation as .pptx or .pptm before applying the navbar animation.");
            }

            string extension = Path.GetExtension(presentation.FullName);
            if (!string.Equals(extension, ".pptx", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".pptm", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Navbar animation currently supports only .pptx or .pptm presentations.");
            }
        }

        private void ApplyAbsoluteScalePatch(PowerPoint.Presentation presentation, int slideIndex, IList<int> targetShapeIds)
        {
            presentation.Save();

            string presentationPath = presentation.FullName;
            string tempRoot = Path.Combine(Path.GetTempPath(), "PPTAssistant", "NavbarAnimationPatch", Guid.NewGuid().ToString("N"));
            string extractedPath = Path.Combine(tempRoot, "extracted");
            string sourceCopyPath = Path.Combine(tempRoot, "source" + Path.GetExtension(presentationPath));
            string patchedArchivePath = Path.Combine(tempRoot, "patched.zip");

            Directory.CreateDirectory(tempRoot);
            File.Copy(presentationPath, sourceCopyPath, true);
            ZipFile.ExtractToDirectory(sourceCopyPath, extractedPath);

            string slideXmlPath = Path.Combine(extractedPath, "ppt", "slides", $"slide{slideIndex}.xml");
            PatchScaleAnimationsInSlideXml(slideXmlPath, targetShapeIds);

            ZipFile.CreateFromDirectory(extractedPath, patchedArchivePath);

            presentation.Close();
            File.Copy(patchedArchivePath, presentationPath, true);

            PowerPoint.Presentation reopenedPresentation = application.Presentations.Open(
                presentationPath,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue);

            ReSelectPatchedShapes(reopenedPresentation, slideIndex, targetShapeIds);
        }

        private static void PatchScaleAnimationsInSlideXml(string slideXmlPath, IList<int> targetShapeIds)
        {
            var document = new XmlDocument();
            document.PreserveWhitespace = true;
            document.Load(slideXmlPath);

            var namespaceManager = new XmlNamespaceManager(document.NameTable);
            namespaceManager.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            for (int index = 0; index < targetShapeIds.Count; index++)
            {
                string shapeId = targetShapeIds[index].ToString(CultureInfo.InvariantCulture);
                XmlNodeList scaleNodes = document.SelectNodes(
                    $"//p:par[p:cTn/p:childTnLst/p:animScale/p:cBhvr/p:tgtEl/p:spTgt[@spid='{shapeId}']]/p:cTn/p:childTnLst/p:animScale",
                    namespaceManager);

                if (scaleNodes == null || scaleNodes.Count < 2)
                {
                    continue;
                }

                SetAbsoluteScaleValue(document, scaleNodes[0], GrowScalePercent, namespaceManager);
                SetAbsoluteScaleValue(document, scaleNodes[1], ShrinkScalePercent, namespaceManager);
            }

            document.Save(slideXmlPath);
        }

        private static void SetAbsoluteScaleValue(XmlDocument document, XmlNode animScaleNode, float targetPercent, XmlNamespaceManager namespaceManager)
        {
            XmlNode byNode = animScaleNode.SelectSingleNode("p:by", namespaceManager);
            if (byNode != null)
            {
                animScaleNode.RemoveChild(byNode);
            }

            XmlNode fromNode = animScaleNode.SelectSingleNode("p:from", namespaceManager);
            if (fromNode != null)
            {
                animScaleNode.RemoveChild(fromNode);
            }

            XmlElement toElement = animScaleNode.SelectSingleNode("p:to", namespaceManager) as XmlElement;
            if (toElement == null)
            {
                toElement = document.CreateElement("p", "to", "http://schemas.openxmlformats.org/presentationml/2006/main");
                animScaleNode.AppendChild(toElement);
            }

            string value = Math.Round(targetPercent * 1000f).ToString(CultureInfo.InvariantCulture);
            toElement.SetAttribute("x", value);
            toElement.SetAttribute("y", value);
        }

        private static void ReSelectPatchedShapes(PowerPoint.Presentation presentation, int slideIndex, IList<int> targetShapeIds)
        {
            PowerPoint.Slide slide = presentation.Slides[slideIndex];
            slide.Select();

            var matchedShapeNames = new List<string>();
            for (int index = 1; index <= slide.Shapes.Count; index++)
            {
                PowerPoint.Shape shape = slide.Shapes[index];
                for (int idIndex = 0; idIndex < targetShapeIds.Count; idIndex++)
                {
                    if (shape.Id == targetShapeIds[idIndex])
                    {
                        matchedShapeNames.Add(shape.Name);
                        break;
                    }
                }
            }

            if (matchedShapeNames.Count == 0)
            {
                return;
            }

            object[] shapeNames = new object[matchedShapeNames.Count];
            for (int index = 0; index < matchedShapeNames.Count; index++)
            {
                shapeNames[index] = matchedShapeNames[index];
            }

            slide.Shapes.Range(shapeNames).Select();
        }
    }
}
