using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Effects
{
    internal sealed class BackgroundImageService
    {
        private const string GeneratedForegroundTag = "PPTAssistantForeground";
        private const int RenderScale = 4;
        private const float BackgroundBlurAmount = 20f;
        private readonly PowerPoint.Application application;

        public BackgroundImageService(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void ApplyFromSelectedPicture()
        {
            PowerPoint.DocumentWindow activeWindow = application.ActiveWindow;
            if (activeWindow == null)
            {
                throw new InvalidOperationException("No active PowerPoint window was found.");
            }

            PowerPoint.Selection selection = activeWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count != 1)
            {
                throw new InvalidOperationException("Select exactly one picture first.");
            }

            PowerPoint.Shape selectedShape = selection.ShapeRange[1];
            if (!IsSupportedPicture(selectedShape))
            {
                throw new InvalidOperationException("Background works only with one selected picture.");
            }

            PowerPoint.Slide slide = selectedShape.Parent as PowerPoint.Slide;
            if (slide == null)
            {
                throw new InvalidOperationException("The selected picture must be on a slide.");
            }

            string tempFolder = Path.Combine(Path.GetTempPath(), "PPTAssistant");
            Directory.CreateDirectory(tempFolder);
            CleanupOldFiles(tempFolder);

            string clearFullPath = Path.Combine(tempFolder, $"clear-full-{Guid.NewGuid():N}.png");
            string blurredPath = Path.Combine(tempFolder, $"blur-{Guid.NewGuid():N}.png");

            float slideWidth = application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = application.ActivePresentation.PageSetup.SlideHeight;
            ExportSlideSizedImage(selectedShape, clearFullPath, slideWidth, slideHeight);
            CreateBlurredCopy(clearFullPath, blurredPath);

            RemoveGeneratedForegrounds(slide);
            selectedShape.Delete();

            PowerPoint.Shape foreground = slide.Shapes.AddPicture(
                clearFullPath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                0f,
                0f,
                slideWidth,
                slideHeight);

            foreground.Name = $"PPTAssistant Foreground {DateTime.Now:HHmmss}";
            foreground.Tags.Add(GeneratedForegroundTag, "1");
            foreground.Locked = MsoTriState.msoTrue;
            foreground.ZOrder(MsoZOrderCmd.msoSendToBack);

            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.UserPicture(blurredPath);
        }

        private static bool IsSupportedPicture(PowerPoint.Shape shape)
        {
            switch (shape.Type)
            {
                case MsoShapeType.msoPicture:
                case MsoShapeType.msoLinkedPicture:
                    return true;
                default:
                    return false;
            }
        }

        private static void RemoveGeneratedForegrounds(PowerPoint.Slide slide)
        {
            for (int index = slide.Shapes.Count; index >= 1; index--)
            {
                PowerPoint.Shape shape = slide.Shapes[index];
                if (shape.Tags[GeneratedForegroundTag] == "1")
                {
                    shape.Delete();
                }
            }
        }

        private static void CleanupOldFiles(string folder)
        {
            foreach (string filePath in Directory.GetFiles(folder, "*.png"))
            {
                try
                {
                    DateTime lastWrite = File.GetLastWriteTimeUtc(filePath);
                    if (lastWrite < DateTime.UtcNow.AddDays(-2))
                    {
                        File.Delete(filePath);
                    }
                }
                catch
                {
                }
            }
        }

        private static void ExportSlideSizedImage(PowerPoint.Shape sourceShape, string outputPath, float slideWidth, float slideHeight)
        {
            int pixelWidth = Math.Max(1, (int)Math.Round(slideWidth * RenderScale));
            int pixelHeight = Math.Max(1, (int)Math.Round(slideHeight * RenderScale));
            sourceShape.Export(
                outputPath,
                PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                pixelWidth,
                pixelHeight,
                PowerPoint.PpExportMode.ppScaleXY);
        }

        private static void CreateBlurredCopy(string sourcePath, string outputPath)
        {
            using (var sourceBitmap = new Bitmap(sourcePath))
            using (var workingBitmap = new Bitmap(sourceBitmap))
            {
                ApplyGaussianBlur(workingBitmap, BackgroundBlurAmount);
                workingBitmap.Save(outputPath, ImageFormat.Png);
            }
        }

        private static void ApplyGaussianBlur(Bitmap bitmap, float blurAmount)
        {
            Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            BitmapData bitmapData = bitmap.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            int stride = bitmapData.Stride;
            int byteCount = stride * bitmap.Height;
            byte[] source = new byte[byteCount];
            byte[] temp = new byte[byteCount];
            byte[] blurred = new byte[byteCount];
            float sigma = Math.Max(0.5f, blurAmount / 2.5f);
            float[] kernel = CreateGaussianKernel(sigma);

            Marshal.Copy(bitmapData.Scan0, source, 0, byteCount);

            HorizontalBlur(source, temp, bitmap.Width, bitmap.Height, stride, kernel);
            VerticalBlur(temp, blurred, bitmap.Width, bitmap.Height, stride, kernel);

            Marshal.Copy(blurred, 0, bitmapData.Scan0, byteCount);
            bitmap.UnlockBits(bitmapData);
        }

        private static void HorizontalBlur(byte[] source, byte[] destination, int width, int height, int stride, float[] kernel)
        {
            int radius = kernel.Length / 2;

            for (int y = 0; y < height; y++)
            {
                int row = y * stride;

                for (int x = 0; x < width; x++)
                {
                    float blue = 0f;
                    float green = 0f;
                    float red = 0f;
                    float alpha = 0f;

                    for (int offset = -radius; offset <= radius; offset++)
                    {
                        int sampleX = Math.Max(0, Math.Min(width - 1, x + offset));
                        int index = row + sampleX * 4;
                        float weight = kernel[offset + radius];
                        blue += source[index] * weight;
                        green += source[index + 1] * weight;
                        red += source[index + 2] * weight;
                        alpha += source[index + 3] * weight;
                    }

                    int destinationIndex = row + x * 4;
                    destination[destinationIndex] = ToByte(blue);
                    destination[destinationIndex + 1] = ToByte(green);
                    destination[destinationIndex + 2] = ToByte(red);
                    destination[destinationIndex + 3] = ToByte(alpha);
                }
            }
        }

        private static void VerticalBlur(byte[] source, byte[] destination, int width, int height, int stride, float[] kernel)
        {
            int radius = kernel.Length / 2;

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    float blue = 0f;
                    float green = 0f;
                    float red = 0f;
                    float alpha = 0f;

                    for (int offset = -radius; offset <= radius; offset++)
                    {
                        int sampleY = Math.Max(0, Math.Min(height - 1, y + offset));
                        int index = sampleY * stride + x * 4;
                        float weight = kernel[offset + radius];
                        blue += source[index] * weight;
                        green += source[index + 1] * weight;
                        red += source[index + 2] * weight;
                        alpha += source[index + 3] * weight;
                    }

                    int destinationIndex = y * stride + x * 4;
                    destination[destinationIndex] = ToByte(blue);
                    destination[destinationIndex + 1] = ToByte(green);
                    destination[destinationIndex + 2] = ToByte(red);
                    destination[destinationIndex + 3] = ToByte(alpha);
                }
            }
        }

        private static float[] CreateGaussianKernel(float sigma)
        {
            int radius = Math.Max(1, (int)Math.Ceiling(sigma * 3f));
            float[] kernel = new float[radius * 2 + 1];
            float sigmaSquaredTimesTwo = 2f * sigma * sigma;
            float totalWeight = 0f;

            for (int offset = -radius; offset <= radius; offset++)
            {
                float weight = (float)Math.Exp(-(offset * offset) / sigmaSquaredTimesTwo);
                kernel[offset + radius] = weight;
                totalWeight += weight;
            }

            for (int index = 0; index < kernel.Length; index++)
            {
                kernel[index] /= totalWeight;
            }

            return kernel;
        }

        private static byte ToByte(float value)
        {
            return (byte)Math.Max(0, Math.Min(255, (int)Math.Round(value)));
        }
    }
}
