using System;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class PowerPointShapeContext
    {
        private PowerPointShapeContext(object owner, PowerPoint.Shapes shapes, bool isTemplateSurface)
        {
            Owner = owner;
            Shapes = shapes;
            IsTemplateSurface = isTemplateSurface;
        }

        public object Owner { get; }

        public PowerPoint.Shapes Shapes { get; }

        public bool IsTemplateSurface { get; }

        public static PowerPointShapeContext FromActiveView(PowerPoint.Application application)
        {
            PowerPoint.DocumentWindow activeWindow = GetActiveWindow(application);

            PowerPointShapeContext context = TryCreateFromSelection(activeWindow.Selection);
            if (context != null)
            {
                return context;
            }

            context = TryCreateFromView(activeWindow);
            if (context != null)
            {
                return context;
            }

            throw new InvalidOperationException("Open a slide, slide master, or layout master before applying this effect.");
        }

        public static PowerPointShapeContext FromShape(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            PowerPointShapeContext context = TryCreate(shape.Parent);
            if (context == null)
            {
                throw new InvalidOperationException("The selected shape must be on a slide, slide master, or layout master.");
            }

            return context;
        }

        public static PowerPoint.Selection GetActiveSelection(PowerPoint.Application application)
        {
            return GetActiveWindow(application).Selection;
        }

        public static void TrySetShapeLocked(PowerPoint.Shape shape, MsoTriState locked)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                dynamic dynamicShape = shape;
                dynamicShape.Locked = locked;
            }
            catch (RuntimeBinderException)
            {
            }
            catch (COMException)
            {
            }
        }

        public PowerPoint.TimeLine GetTimeLine()
        {
            try
            {
                dynamic owner = Owner;
                return (PowerPoint.TimeLine)owner.TimeLine;
            }
            catch (RuntimeBinderException)
            {
            }
            catch (InvalidCastException)
            {
            }

            throw new InvalidOperationException("PowerPoint does not expose an animation timeline for the current editing surface.");
        }

        public void SetBackgroundPicture(string picturePath)
        {
            try
            {
                dynamic owner = Owner;
                owner.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            catch (RuntimeBinderException)
            {
            }

            try
            {
                dynamic owner = Owner;
                owner.Background.Fill.UserPicture(picturePath);
            }
            catch (RuntimeBinderException)
            {
                throw new InvalidOperationException("PowerPoint does not expose a background fill for the current editing surface.");
            }
        }

        public void ClearShapesIfTemplateSurface()
        {
            if (!IsTemplateSurface)
            {
                return;
            }

            ClearShapes();
        }

        public void ClearShapes()
        {
            for (int index = Shapes.Count; index >= 1; index--)
            {
                PowerPoint.Shape shape = Shapes[index];
                try
                {
                    TrySetShapeLocked(shape, MsoTriState.msoFalse);
                }
                catch (RuntimeBinderException)
                {
                }

                shape.Delete();
            }
        }

        private static PowerPoint.DocumentWindow GetActiveWindow(PowerPoint.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException(nameof(application));
            }

            PowerPoint.DocumentWindow activeWindow = application.ActiveWindow;
            if (activeWindow == null)
            {
                throw new InvalidOperationException("No active PowerPoint window was found.");
            }

            return activeWindow;
        }

        private static PowerPointShapeContext TryCreateFromView(PowerPoint.DocumentWindow activeWindow)
        {
            try
            {
                return TryCreate(activeWindow.View.Slide, IsTemplateView(activeWindow.ViewType));
            }
            catch (RuntimeBinderException)
            {
            }
            catch (InvalidCastException)
            {
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static PowerPointShapeContext TryCreateFromSelection(PowerPoint.Selection selection)
        {
            try
            {
                if (selection != null &&
                    selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    selection.ShapeRange.Count > 0)
                {
                    return TryCreate(selection.ShapeRange[1].Parent);
                }
            }
            catch (RuntimeBinderException)
            {
            }
            catch (InvalidCastException)
            {
            }

            return null;
        }

        private static PowerPointShapeContext TryCreate(object owner, bool templateSurfaceOverride = false)
        {
            if (owner == null)
            {
                return null;
            }

            try
            {
                dynamic dynamicOwner = owner;
                PowerPoint.Shapes shapes = (PowerPoint.Shapes)dynamicOwner.Shapes;
                return shapes == null ? null : new PowerPointShapeContext(owner, shapes, templateSurfaceOverride || IsTemplateOwner(owner));
            }
            catch (RuntimeBinderException)
            {
            }
            catch (InvalidCastException)
            {
            }

            return null;
        }

        private static bool IsTemplateOwner(object owner)
        {
            return owner is PowerPoint.Master || owner is PowerPoint.CustomLayout;
        }

        private static bool IsTemplateView(PowerPoint.PpViewType viewType)
        {
            switch (viewType)
            {
                case PowerPoint.PpViewType.ppViewSlideMaster:
                case PowerPoint.PpViewType.ppViewTitleMaster:
                case PowerPoint.PpViewType.ppViewHandoutMaster:
                case PowerPoint.PpViewType.ppViewNotesMaster:
                case PowerPoint.PpViewType.ppViewMasterThumbnails:
                    return true;
                default:
                    return false;
            }
        }
    }
}
