using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace ExportExcel
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool bCreateMutex;
            using (System.Threading.Mutex kMutex = new System.Threading.Mutex(true, Application.ProductName, out bCreateMutex))
            {
                if (!bCreateMutex)
                {
                    MessageBox.Show("系统已经在运行,请勿重复开启");
                    return;
                }
            }

            Process[] process = System.Diagnostics.Process.GetProcesses();

            for (int i = 0; i < process.Length; i++)
            {
                if (process[i].ProcessName.Length>3)
                {
                    if (process[i].ProcessName.Substring(0, 3) == "WPS")
                    {

                    }
                    
                }                
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new WPSExportForm());
        }
    }
}
