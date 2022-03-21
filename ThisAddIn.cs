//using System;
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
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace xlladdin
{
    /*
        

        public enum Operation
        {
            Add = 1,
            Remove = 2,
            New = 3
        }
        public void Call(Operation op, string name)
        {
            Application.ExecuteExcel4Macro($"ADDIN.MANAGER(op {name}");
        }

     }
    */
    public partial class ThisAddIn
    {
         
        // GitHub URLs
        private readonly string AddInURL = @"https://github.com/xlladdins/";
        private readonly string RawURL = @"https://raw.githubusercontent.com/xlladdins/";
        
        // Excel template directory of user
        private readonly string AddInDir = Environment.GetEnvironmentVariable("AppData") + @"\Microsoft\AddIns\";

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

        ///
        /// https://xlladdins.github.io/Excel4Macros/addin.manager.html
        /// Excel moves AIM value to OPT and adds OPEN<n+1> subkey to load at startup.
        ///

        // Registry entries used by add-in manager.
        private string OFFICE = "Software\\Microsoft\\Office\\";
        private string AIM = "\\Excel\\Add-in Manager";
        private string OPT = "\\Excel\\Options";

        // Excel version
        private string Version()
        {
            return Application.ExecuteExcel4Macro("GET.WORKSPACE(2)");
        }

        // Adds an add-in to the working set using the descriptive name in the Add-Ins dialog box.
        private dynamic Add(string name)
		{
            return Application.ExecuteExcel4Macro($"ADDIN.MANAGER(1, \"{name}\")");
		}
        // Remove an add-in from the working set using the descriptive name in the Add-Ins dialog box.
        private dynamic Remove(string name)
        {
            return Application.ExecuteExcel4Macro($"ADDIN.MANAGER(2, \"{name}\")");
        }
        // Adds a new add-in to the working set using the full file name to in the Add-Ins dialog box.
        private dynamic New(string file)
        {
            return Application.ExecuteExcel4Macro($"ADDIN.MANAGER(3, \"{file}\")");
        }

        // Full path if descriptive name matches AIM value.
        private bool Known(string name)
        {
            RegistryKey aim = Registry.CurrentUser.OpenSubKey(OFFICE + Version() + AIM);

            foreach (string key in aim.GetSubKeyNames())
            {
                if (aim.GetValue(key).ToString().Contains(name))
                {
                    return true;
                }
            }

            return false;
        }

        // OPT key OPEN<n> contains name if loaded
        bool Loaded(string name)
        {
            RegistryKey opt = Registry.CurrentUser.OpenSubKey(OFFICE + Version() + OPT);

            foreach (string key in opt.GetSubKeyNames())
            {
                if (key.StartsWith("OPEN") && opt.GetValue(key).ToString().Contains(name))
                { 
                    return true;
                }
            }

            return false;
        }
        private bool Unregister(string module)
        {
            return true == Application.ExecuteExcel4Macro($"UNREGISTER(\"{module}\")");
        }
        private void Register(string module)
        {
            Application.ExecuteExcel4Macro($"OPEN(\"{module}\")");
        }

        /// <summary>
        /// Download file from url to dir if newer than date.
        /// </summary>
        private void Download(string url, string dir, string file, DateTime date)
        {
            string name = Path.GetFileNameWithoutExtension(file);
            bool exists = File.Exists(dir + file);
            bool newer = exists && File.GetLastWriteTime(dir + file) < date;
            bool download = !exists || newer;

            if (exists && newer)
            {
                download = DialogResult.Yes ==
                    MessageBox.Show(
                        $"Download newer version of {file}?", "Download",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            if (download)
            {
                //bool registered = Unregister(dir + file);
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                WebClient webClient = new WebClient();
                try
                {
                    using (Stream istream = webClient.OpenRead(url + file))
                    {
                        using (Stream ostream = File.OpenWrite(dir + file))
                        {
                            istream.CopyTo(ostream);
                        }
                        New(dir + file);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Download all known addins
        private void Addins(string url, string files)
        {
            string bits = Bits();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            WebClient webClient = new WebClient();

            // text file of available add-ins
            using (Stream istream = webClient.OpenRead(url + files + "?ticks=" + DateTime.Now.Ticks.ToString()))
            {
                using (StreamReader sr = new StreamReader(istream))
                {
                    while (sr.Peek() != -1)
                    {
                        string[] filedate = sr.ReadLine().Split(' ');
                        string file = filedate[0];
                        DateTime date = DateTime.Parse(filedate[1]);
                        string xll = file + ".xll";
                        // add/remove files
                        //Task.Factory.StartNew(() => { 
                        Download(AddInURL + file + @"/raw/master/x" + bits + @"/", AddInDir, xll, date);
                        //});
                    }
                }
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Addins(RawURL, @"xlladdin/master/xlladdins.txt");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //MessageBox.Show("shutdown");
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
