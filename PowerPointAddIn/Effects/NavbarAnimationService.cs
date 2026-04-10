using System;
using System.Globalization;
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
        private const float BaseScalePercent = 100f;
        private readonly PowerPoint.Application application;

        public NavbarAnimationService(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public int ApplyToSelection()
        {
            PowerPoint.ShapeRange shapeRange = GetSelectedShapeRange();
            int changedCount = 0;

            for (int index = 1; index <= shapeRange.Count; index++)
            {
                ApplyToShape(shapeRange[index]);
                changedCount++;
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
            growEffect.Behaviors[1].ScaleEffect.ByX = GrowScalePercent - BaseScalePercent;
            growEffect.Behaviors[1].ScaleEffect.ByY = GrowScalePercent - BaseScalePercent;

            PowerPoint.Effect shrinkEffect = sequence.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink);
            shrinkEffect.Timing.Duration = ShrinkDurationSeconds;
            shrinkEffect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            shrinkEffect.Timing.TriggerDelayTime = ShrinkDelaySeconds;
            shrinkEffect.Timing.Accelerate = 0f;
            shrinkEffect.Timing.Decelerate = 0f;
            shrinkEffect.Behaviors[1].ScaleEffect.ByX = ShrinkScalePercent - BaseScalePercent;
            shrinkEffect.Behaviors[1].ScaleEffect.ByY = ShrinkScalePercent - BaseScalePercent;
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
    }
}
