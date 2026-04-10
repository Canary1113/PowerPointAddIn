using System;
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
            float slideWidth = application.ActivePresentation.PageSetup.SlideWidth;
            float moveDistancePercent = shape.Width / slideWidth * 100f;

            PowerPoint.Effect moveEffect = sequence.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectCustom,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious,
                -1);
            moveEffect.Timing.Duration = MoveDurationSeconds;
            PowerPoint.AnimationBehavior motionBehavior = moveEffect.Behaviors.Add(
                PowerPoint.MsoAnimType.msoAnimTypeMotion,
                1);
            motionBehavior.MotionEffect.ByX = moveDistancePercent;
            motionBehavior.MotionEffect.ByY = 0f;

            PowerPoint.Effect growEffect = sequence.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectCustom,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious,
                -1);
            growEffect.Timing.Duration = GrowDurationSeconds;
            PowerPoint.AnimationBehavior growBehavior = growEffect.Behaviors.Add(
                PowerPoint.MsoAnimType.msoAnimTypeScale,
                1);
            growBehavior.ScaleEffect.FromX = 100f;
            growBehavior.ScaleEffect.FromY = 100f;
            growBehavior.ScaleEffect.ToX = GrowScalePercent;
            growBehavior.ScaleEffect.ToY = GrowScalePercent;

            PowerPoint.Effect shrinkEffect = sequence.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectCustom,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious,
                -1);
            shrinkEffect.Timing.Duration = ShrinkDurationSeconds;
            shrinkEffect.Timing.TriggerDelayTime = ShrinkDelaySeconds;
            PowerPoint.AnimationBehavior shrinkBehavior = shrinkEffect.Behaviors.Add(
                PowerPoint.MsoAnimType.msoAnimTypeScale,
                1);
            shrinkBehavior.ScaleEffect.FromX = GrowScalePercent;
            shrinkBehavior.ScaleEffect.FromY = GrowScalePercent;
            shrinkBehavior.ScaleEffect.ToX = ShrinkScalePercent;
            shrinkBehavior.ScaleEffect.ToY = ShrinkScalePercent;
        }
    }
}
