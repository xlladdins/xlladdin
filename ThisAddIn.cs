﻿//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Xml.Linq;
//using Excel = Microsoft.Office.Interop.Excel;
//using Office = Microsoft.Office.Core;
//using Microsoft.Office.Tools.Excel;
using Microsoft.Win32;
using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace xlladdin
{
    public partial class ThisAddIn
    {
        private static string AddInURL = "https://xlladdins.com/addins/";
        // Excel template directory of user
        private static string AddInDir = Environment.GetEnvironmentVariable("AppData") + @"\Microsoft\AddIns\";

        [DllImport("kernel32")]
        public extern static IntPtr LoadLibrary(string librayName);

        [DllImport("kernel32.dll", EntryPoint = "FreeLibrary")]
        static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32", CharSet = CharSet.Ansi)]
        public extern static IntPtr GetProcAddress(IntPtr hwnd, string procedureName);

        public enum BinaryType : uint
        {
            SCS_32BIT_BINARY = 0, // A 32-bit Windows-based application
            SCS_64BIT_BINARY = 6, // A 64-bit Windows-based application.
            SCS_DOS_BINARY = 1, // An MS-DOS – based application
            SCS_OS216_BINARY = 5, // A 16-bit OS/2-based application
            SCS_PIF_BINARY = 3, // A PIF file that executes an MS-DOS – based application
            SCS_POSIX_BINARY = 4, // A POSIX – based application
            SCS_WOW_BINARY = 2  // A 16-bit Windows-based application 
        }
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool GetBinaryType(string lpApplicationName, out BinaryType lpBinaryType);

        private static BinaryType ExcelBinaryType()
        {
            BinaryType type;

            string app_path = @"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe";
            using (var key = Registry.LocalMachine.OpenSubKey(app_path))
            {
                var excel = key.GetValue("");
                GetBinaryType(excel.ToString(), out type);
            }

            return type;
        }
        /// <summary>
        /// Determine if Excel is 32 or 64-bit.
        /// </summary>
        /// <returns>Either "32" or "64".</returns>
        /// <exception cref="Exception"></exception>
        private static string Bits()
        {
            string bits = null;

            switch (ExcelBinaryType())
            {
                case BinaryType.SCS_32BIT_BINARY:
                    bits = "32";
                    break;
                case BinaryType.SCS_64BIT_BINARY:
                    bits = "64";
                    break;
            }

            if (bits == null)
            {
                throw new Exception("unable to determine Excel bitness");
            }

            return bits;
        }

        /// <summary>
        /// Download file from url
        /// </summary>
        private static void Download(string url, string dir)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            WebClient webClient = new WebClient();

            using (Stream istream = webClient.OpenRead(url + $"xlladdin{Bits()}.xll"))
            {
                using (Stream ostream = File.OpenWrite(dir + "xlladdins.xll"))
                {
                    istream.CopyTo(ostream);
                }
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Download(AddInURL, AddInDir);
                Application.RegisterXLL(AddInDir + "xlladdins.xll");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            MessageBox.Show("shudown");
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}