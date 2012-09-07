namespace Resx2Xls
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Drawing;
    using System.Drawing.Drawing2D;
    using System.Drawing.Imaging;
    using System.IO;

    public static class ImageHelper
    {
        public enum CodecType
        {
            PNG, JPG,
        }

        public static void ResizeImage(string sourcePath, string targetPath, int newWidth, int newHeight, CodecType codec)
        {
            Bitmap newImage = new Bitmap(newWidth, newHeight);
            using (Graphics gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(Bitmap.FromFile(sourcePath), new Rectangle(0, 0, newWidth, newHeight));
            }

            ImageCodecInfo codecInfo;
            EncoderParameters encoderParameters;

            switch(codec)
            {
                case CodecType.JPG:
                    codecInfo = GetEncoder(ImageFormat.Jpeg);
                    encoderParameters = new EncoderParameters(1);
                    encoderParameters.Param[0] = new EncoderParameter(Encoder.Quality, 75L);
                    break;

                case CodecType.PNG:
                    codecInfo = GetEncoder(ImageFormat.Png);
                    encoderParameters = new EncoderParameters(0);
                    break;

                default:
                    throw new InvalidOperationException("Unknown codec");
            }
            
            newImage.Save(targetPath, codecInfo, encoderParameters);
        }

        static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            return codecs.Single(codec => codec.FormatID == format.Guid);
        }

        internal static bool GetScaledImage(string sourcePath, int maxWidth, out int width, out int height, out string targetPath)
        {
            using( var source = Bitmap.FromFile(sourcePath))
            {
                if (source.Width < maxWidth)
                {
                    targetPath = sourcePath;
                    width = source.Width;
                    height = source.Height;
                    return false;
                }
                else
                {
                    targetPath = Path.GetTempFileName();
                    width = maxWidth;
                    height = maxWidth * source.Height / source.Width;
                    ResizeImage(sourcePath, targetPath, width, height, CodecType.JPG);
                    return true;
                }
            }
        }
    }
}
