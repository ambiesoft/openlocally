using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Trinet.Networking;
using NDesk.Options;

namespace openlocally
{
    static class Program
    {
        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPTStr)] string localName,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
            ref int length);

        static void showVersion()
        {
            MessageBox.Show(Application.ProductName +
    " ver" +
    Ambiesoft.AmbLib.getAssemblyVersion(Assembly.GetExecutingAssembly(), 3),
    Application.ProductName,
    MessageBoxButtons.OK,
    MessageBoxIcon.Information);
        }

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            string program = string.Empty;


            List<string> extraCommandLineArgs = new List<string>();

            var p = new OptionSet()
                .Add("v|version", dummy => { showVersion(); Environment.Exit(0); })
                .Add("p=|program=", prog => { if (prog != null) program = prog; })
                ;

            try
            {
                extraCommandLineArgs = p.Parse(args);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(0);
            }

            if (extraCommandLineArgs.Count == 0)
                exitProgram(Properties.Resources.NO_ARGUMENTS);
            if (extraCommandLineArgs.Count != 1)
                exitProgram(Properties.Resources.NO_ARGUMENTS);
            string inputpath = extraCommandLineArgs[0];

            string fullpath = Path.GetFullPath(inputpath);
            string localpath;
            if (Char.IsLetter(fullpath[0]) &&
                fullpath[1] == ':' &&
                fullpath[2] == '\\')
            {
                // First assume fullpath is Network Drive
                var sb = new StringBuilder(512);
                var size = sb.Capacity;
                var error = WNetGetConnection(fullpath.Substring(0, 2), sb, ref size);
                if (error == 0)
                {
                    // success, fullpath is Network Drive
                    var networkpath = sb.ToString();
                    if (networkpath.Length == 0)
                        exitProgram("Network paht is null");
                    networkpath += fullpath.Substring(2);

                    localpath = GetLocalPath(networkpath);
                }
                else
                {
                    // normal drive
                    // exitProgram((new Win32Exception(error, "WNetGetConnection failed").ToString()));
                    localpath = fullpath;
                }
            }
            else
            {
                localpath = GetLocalPath(inputpath);
            }

            if (localpath == null || localpath.Length == 0)
                exitProgram(Properties.Resources.LOCAL_PATH_NOT_FOUND);

            openInExplorer(localpath,program);
            
            
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new FormMain());
        }

        // https://stackoverflow.com/a/696144
        static void openInExplorer(string filePath, string program)
        {
            if (!File.Exists(filePath) && !Directory.Exists(filePath))
            {
                exitProgram(string.Format(Properties.Resources.FILE_OR_FOLDER_NOT_EXIST, filePath));
                return;
            }

            try
            {
                if (string.IsNullOrEmpty(program))
                {
                    // combine the arguments together
                    // it doesn't matter if there is a space after ','
                    string argument = "/select, \"" + filePath + "\"";
                    System.Diagnostics.Process.Start("explorer.exe", argument);
                }
                else
                {
                    if (!File.Exists(program))
                    {
                        exitProgram(string.Format(Properties.Resources.PROGRAM_NOT_EXIST, program));
                    }
                    System.Diagnostics.Process.Start(program, Ambiesoft.AmbLib.doubleQuoteIfSpace(filePath));
                }
            }
            catch(Exception ex)
            {
                exitProgram(ex.Message);
            }
        }
        static string getServer(string netfile)
        {
            if (netfile.Length < 3)
                return null;
            if (netfile[0] != '\\' || netfile[1] != '\\')
                return null;
            int firstsep = netfile.IndexOf('\\', 2);
            if (firstsep <= 0)
                return null;

            return netfile.Substring(2, firstsep - 2);
        }
        static string getServerAndShare(string netfile)
        {
            if (netfile.Length < 3)
                return null;
            if (netfile[0] != '\\' || netfile[1] != '\\')
                return null;
            int firstsep = netfile.IndexOf('\\', 2);
            if (firstsep <= 0)
                return null;

            int secondsep = netfile.IndexOf('\\', firstsep+1);
            if (secondsep <= 0)
                return null;

            return netfile.Substring(0, secondsep);
        }
        static void exitProgram(string error)
        {
            MessageBox.Show(error, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            Environment.Exit(1);
        }

        // https://stackoverflow.com/a/8809437
        static string ReplaceFirst(string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
        static string GetLocalPath(string netfile)
        {
            string server = getServer(netfile);
            if (server == null || server.ToLower() != Environment.MachineName.ToLower())
                exitProgram(Properties.Resources.SERVER_NULL);
            ShareCollection shi = ShareCollection.LocalShares;
            if (shi == null)
                exitProgram(Properties.Resources.SHARECOLLECTION_NULL);

            string serverandshare = getServerAndShare(netfile);
            if (serverandshare == null)
                exitProgram(Properties.Resources.SERVERANDSHARE_NULL);
            foreach (Share si in shi)
            {
                if (si.ShareType==ShareType.Disk && si.IsFileSystem)
                {
                    if(serverandshare.ToLower()==si.ToString().ToLower())
                    {
                        return ReplaceFirst(netfile, serverandshare, si.Path);
                    }
                }
            }
            return null;
        }
    }
}
