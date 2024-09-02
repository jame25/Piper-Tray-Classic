using System;
using System.Threading;
using System.Windows.Forms;

namespace PiperTrayClassic
{
    static class Program
    {
        static Mutex mutex = new Mutex(true, "{8F6F0AC4-B9A1-45fd-A8CF-72F04E6BDE8F}");

        [STAThread]
        static void Main()
        {
            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                try
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new PiperTrayApp());
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
            else
            {
                MessageBox.Show("Another instance of Piper Tray Classic is already running.", "Piper Tray Classic", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
