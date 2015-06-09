//-----------------------------------------------------------------------
// <copyright file="PushPin.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Class representing the PushPin and responsible for providing the BitmapImage for
    /// all the pushpin for rendering them in custom task pane.
    /// </summary>
    internal class PushPin
    {
        /// <summary>
        /// Image having all the pushpins.
        /// </summary>
        private static Bitmap pushPinsImage = Properties.Resources.Pins;

        /// <summary>
        /// Caching of the BitmapImage of all the push pins.
        /// </summary>
        private static Dictionary<int, BitmapImage> pinBitmapImageCache = new Dictionary<int, BitmapImage>();

        /// <summary>
        /// Gets the count of pins available in the PushPin image.
        /// </summary>
        internal static int PinCount
        {
            get
            {
                return 348;
            }
        }

        /// <summary>
        /// Gets the BitmapImage of the specified push pin.
        /// </summary>
        /// <param name="pinId">Pushpin id</param>
        /// <returns>BitmapImage of the specified push pin</returns>
        internal static BitmapImage GetPushPinBitmapImage(int pinId)
        {
            BitmapImage bitmapImage = null;

            if (pinBitmapImageCache.ContainsKey(pinId))
            {
                bitmapImage = pinBitmapImageCache[pinId];
            }
            else
            {
                using (System.Drawing.Bitmap bitmap = PushPin.GetPushPinBitmap(pinId))
                {
                    using (MemoryStream bitmapStream = new MemoryStream())
                    {
                        bitmap.Save(bitmapStream, System.Drawing.Imaging.ImageFormat.Png);
                        bitmapImage = new System.Windows.Media.Imaging.BitmapImage();

                        bitmapImage.BeginInit();
                        bitmapImage.StreamSource = new MemoryStream(bitmapStream.ToArray());
                        bitmapImage.CreateOptions = BitmapCreateOptions.None;
                        bitmapImage.CacheOption = BitmapCacheOption.Default;
                        bitmapImage.EndInit();

                        pinBitmapImageCache.Add(pinId, bitmapImage);
                    }
                }
            }

            return bitmapImage;
        }

        /// <summary>
        /// Gets the Bitmap of the specified push pin.
        /// </summary>
        /// <param name="pinId">Pushpin id</param>
        /// <returns>Bitmap of the specified push pin</returns>
        private static Bitmap GetPushPinBitmap(int pinId)
        {
            Bitmap bmp = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                int row = pinId / 16;
                int col = pinId % 16;
                graphics.DrawImage(pushPinsImage, new Rectangle(0, 0, 32, 32), (col * 32), (row * 32), 32, 32, GraphicsUnit.Pixel);

                graphics.Flush();
            }

            return bmp;
        }
    }
}
