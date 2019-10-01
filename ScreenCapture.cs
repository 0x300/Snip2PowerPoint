using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Snip2PowerPoint
{
    // This class is responsible for creating the capture overlay (canvas)
    // and capturing an image from the user's selection 
    public static class ScreenCapture
    {
        public static Bitmap GetSnapshotBitmap()
        {
            Rectangle snapshotRect = Screen.GetBounds(Point.Empty);
            using (Canvas canvas = new Canvas())
            {
                // TODO: figure out why this has to be set here and not in the canvas InitializeComponent() call
                canvas.TopMost = true;
                if (canvas.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    snapshotRect = canvas.GetRectangle();
                }
            }

            using (Image image = new Bitmap(snapshotRect.Width, snapshotRect.Height))
            {
                using (Graphics graphics = Graphics.FromImage(image))
                {
                    graphics.CopyFromScreen(new Point(snapshotRect.Left, snapshotRect.Top), Point.Empty, snapshotRect.Size);
                }
                return new Bitmap(SetBorder(image, Color.Black, 1));
            }
        }

        private static Image SetBorder(Image srcImg, Color color, int width)
        {
            // Create a copy of the image and graphics context
            Image dstImg = srcImg.Clone() as Image;
            Graphics g = Graphics.FromImage(dstImg);
            
            // Create the pen
            Pen pBorder = new Pen(color, width)
            {
                Alignment = PenAlignment.Center
            };

            // Draw
            g.DrawRectangle(pBorder, 0, 0, dstImg.Width - 1, dstImg.Height - 1);

            // Clean up
            pBorder.Dispose();
            g.Save();
            g.Dispose();

            // Return
            return dstImg;
        }
    }
}
