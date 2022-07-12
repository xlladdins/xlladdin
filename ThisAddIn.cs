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
//using System.Threading.Tasks;
using System.Windows.Forms;
using WinSCP;

namespace xlladdin
{
    public partial class ThisAddIn
    {
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

        // Excel version
        private string Version()
        {
            return Application.ExecuteExcel4Macro("GET.WORKSPACE(2)");
        }

        /// <summary>
        /// Download if newer remote version or local version does not exist
        /// </summary>
        /// <param name="fileInfo"></param>
        private void Update(Session session, RemoteFileInfo fileInfo)
        {
            string fileName = AddInDir + fileInfo.Name;
            bool exists = File.Exists(fileName);
            bool newer = exists && File.GetLastWriteTime(fileName) < fileInfo.LastWriteTime;
 
            if (newer)
            {
                var addIn = Application.AddIns2[Path.GetFileNameWithoutExtension(fileInfo.Name)];
                bool installed = addIn.Installed;
                // temporarily unload
                if (installed)
                {
                    addIn.Installed = false;
                }
                session.GetFileToDirectory(fileInfo.FullName, AddInDir, true, null);
                // reload
                if (installed)
                {
                    addIn.Installed = true;
                }
            }
            else if (!exists)
            {
                session.GetFileToDirectory(fileInfo.FullName, AddInDir, true, null);
                Application.AddIns.Add(AddInDir + fileInfo.Name);
                var addIn = Application.AddIns2[Path.GetFileNameWithoutExtension(fileInfo.Name)];
                addIn.Installed = true;
            }
        }

        /// <summary>
        /// Addin files for appropriate Excel bitness.
        /// </summary>
        private void SyncAddins()
        {
            string bits = Bits();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            SessionOptions sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Webdav,
                HostName = "xlladdins.com",
                RootPath = "/addins",
                UserName = "kal",
                Password = "wo3deameh"
            };
            using (Session session = new Session())
            {
                session.Open(sessionOptions);
                // all files in bits directory
                RemoteDirectoryInfo directory = session.ListDirectory(bits);
                foreach (RemoteFileInfo fileInfo in directory.Files)
                {
                    // Can't remove temporary ~$ WebDAV files.
                    if (!fileInfo.IsDirectory && !fileInfo.Name.StartsWith(@"~$"))
                    {
                        Update(session, fileInfo);
                    }
                }

            }
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                SyncAddins();
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
