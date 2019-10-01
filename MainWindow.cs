using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Snip2PowerPoint
{
    public partial class MainWindow : Form
    {
        // Hotkey constants
        private const int MOD_ALT = 0x0001;
        private const int MOD_CONTROL = 0x0002;
        private const int MOD_SHIFT = 0x0004;
        private const int TAKE_SNAP_HOTKEY_ID = 1;

        // TODO: these probably don't need to be globals
        private int snapCount;
        private List<Bitmap> snaps;
        private bool enableSnaps = true;

        public MainWindow()
        {
            InitializeComponent();

            snapCount = 0;
            snaps = new List<Bitmap>();

            // Modifier keys codes: Alt = 1, Ctrl = 2, Shift = 4, Win = 8
            // Compute the addition of each combination of the keys you want to be pressed
            // ALT+CTRL = 1 + 2 = 3 , CTRL+SHIFT = 2 + 4 = 6...
            Win32.RegisterHotKey(this.Handle, TAKE_SNAP_HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (int)Keys.Z);
        }

        // Process system messages (like hotkey presses)
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == TAKE_SNAP_HOTKEY_ID)
            {
                if(enableSnaps)
                {
                    enableSnaps = false;
                    TakeSnap();
                    enableSnaps = true;
                }
            }

            base.WndProc(ref m);
        }

        private void takeSnapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TakeSnap();
        }

        private void TakeSnap()
        {
            var snap = ScreenCapture.GetSnapShot();
            snaps.Add(snap);
            AddToPreview(snap);
            Utility.SaveToOpenPowerPoint(snap);
        }

        private void AddToPreview(Bitmap snap)
        {
            imageList1.Images.Add(ResizeImage(snap, 150, 150));
            listView1.Items.Add(new ListViewItem("" + (++snapCount), imageList1.Images.Count - 1)).EnsureVisible();            
        }

        static Image ResizeImage(Image imgPhoto, int Width, int Height)
        {
            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)Width / (float)sourceWidth);
            nPercentH = ((float)Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = System.Convert.ToInt16((Width -
                              (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = System.Convert.ToInt16((Height -
                              (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap bmPhoto = new Bitmap(Width, Height, PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution, imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.White);
            grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new Rectangle(destX, destY, destWidth, destHeight),
                new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            return bmPhoto;
        }

        private void saveAsImagesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utility.SaveAsImages(snaps);
        }

        private void MainWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            Win32.UnregisterHotKey(this.Handle, TAKE_SNAP_HOTKEY_ID);
        }
    }
}