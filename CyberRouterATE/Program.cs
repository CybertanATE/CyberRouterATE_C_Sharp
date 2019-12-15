using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CyberRouterATE
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool result;
            var mutex = new System.Threading.Mutex(true, "57991c61-d884-4806-a22e-2884d32f43ab", out result);

            if (!result)
            {
                MessageBox.Show("Another CyberRouterATE is already running.", "Warning");
                return;
            }


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new RouterTestMain());
        }
    }
}
