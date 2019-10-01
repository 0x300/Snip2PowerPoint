using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Diagnostics;
using System;

namespace Snip2PowerPoint
{
    class Utility
    {
        public static void SaveAsImages(List<Bitmap> images)
        {
            ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);
            System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;
            EncoderParameters myEncoderParameters = new EncoderParameters(1);
            EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 100L);
            myEncoderParameters.Param[0] = myEncoderParameter;

            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    int count = 1;
                    foreach (Bitmap img in images)
                    {
                        img.Save(dialog.SelectedPath + "\\" + (count++) + ".jpg", jgpEncoder, myEncoderParameters);
                    }
                }
            }
        }

        public static void SaveToOpenPowerPoint(Bitmap image)
        {
            Process[] powerpoints = Process.GetProcessesByName("POWERPNT");

            if (powerpoints.Length < 1)
            {
                // Skip saving to powerpoint if not running..
                return;
            }

            Clipboard.SetImage(image);

            IntPtr savedActiveWindow = Win32.GetForegroundWindow();

            Win32.SwitchToThisWindow(powerpoints[0].MainWindowHandle, true);

            SendKeys.Send("^{m}");
            System.Threading.Thread.Sleep(100);
            SendKeys.Send("^{v}");

            Win32.SwitchToThisWindow(savedActiveWindow, true);
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();

            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }
    }    
}
