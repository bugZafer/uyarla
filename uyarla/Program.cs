using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace uyarla
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool isAppRunning;

            Mutex mutex = new Mutex(true, "{EEDD2892-7666-4D82-AB02-9D12D1C6A54E}", out isAppRunning);

            if (!isAppRunning)
            {
                MessageBox.Show("Uygulama zaten çalışıyor!");
                return;
            }
            Application.Run(new fompk());

        }
    }
}
