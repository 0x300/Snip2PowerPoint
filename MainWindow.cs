using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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

        private int snapCount = 0;
        private List<Bitmap> snaps = new List<Bitmap>();
        private bool snapInProgress = false;

        public MainWindow()
        {
            InitializeComponent();

            // Modifier keys codes: Alt = 1, Ctrl = 2, Shift = 4, Win = 8
            // Compute the addition of each combination of the keys to be pressed
            // ALT+CTRL = 1 + 2 = 3 , CTRL+SHIFT = 2 + 4 = 6...
            Win32.RegisterHotKey(this.Handle, TAKE_SNAP_HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (int)Keys.Z);
        }

        // System message handler to detect screenshot hotkey
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == TAKE_SNAP_HOTKEY_ID)
            {
                // Return if already taking a screenshot and the hotkey gets hit again
                if (snapInProgress) return;
                
                snapInProgress = true;
                TakeSnap();
                snapInProgress = false;
            }

            base.WndProc(ref m);
        }

        private void TakeSnap()
        {
            Bitmap snap = ScreenCapture.GetSnapshotBitmap();
            snaps.Add(snap);
            AddToPreview(snap);
            Utility.SaveToOpenPowerPoint(snap);
        }

        private void AddToPreview(Bitmap snap)
        {
            imageList1.Images.Add(Utility.ResizeImage(snap, 150, 150));
            listView1.Items.Add(new ListViewItem("" + (++snapCount), imageList1.Images.Count - 1)).EnsureVisible();            
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