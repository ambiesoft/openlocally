using NDesk.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Trinet.Networking;

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
            MessageBox.Show(
                Application.ProductName + " ver" + Ambiesoft.AmbLib.getAssemblyVersion(Assembly.GetExecutingAssembly(), 3),
                Application.ProductName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string program = string.Empty;
            bool openRemote = false;

            List<string> extraCommandLineArgs = new List<string>();

            var p = new OptionSet()
                .Add("v|version", dummy => { showVersion(); Environment.Exit(0); })
                .Add("p=|program=", prog => { if (prog != null) program = prog; })
                .Add("o|openremote", dummy => { openRemote = true; })
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
                exitProgram(Properties.Resources.MORE_THAN_1_ARGUMENTS);
            string inputpath = extraCommandLineArgs[0];

            string fullpath = Path.GetFullPath(inputpath);
            string localpath;
            bool changedFromNetworkDrive = false;
            bool originallyLocal = false;
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
                    changedFromNetworkDrive = true;
                }
                else
                {
                    // normal drive
                    // exitProgram((new Win32Exception(error, "WNetGetConnection failed").ToString()));
                    originallyLocal = true;
                    localpath = fullpath;
                }
            }
            else
            {
                // UNC
                localpath = GetLocalPath(inputpath);
                changedFromNetworkDrive = true;
                if (localpath == null)
                {
                    if (openRemote)
                    {
                        openInExplorer(inputpath, program);
                        return;
                    }
                    // exitProgram(string.Format(Properties.Resources.PATH_NOT_HOSTED, netfile));
                }
            }

            if (localpath == null || localpath.Length == 0)
                exitProgram(string.Format(
                    Properties.Resources.LOCAL_PATH_NOT_FOUND,
                    inputpath,
                    Environment.MachineName));

            if ((changedFromNetworkDrive || originallyLocal)
                && !Directory.Exists(localpath)
                && MessageBox.Show(
                    String.Format(
                        (originallyLocal ?
                        Properties.Resources.PATH_IS_ORIGINALLY_LOCAL :
                        Properties.Resources.PATH_RESOLVED_DO_YOU_WANT_TO_OPEN), localpath) +
                        " " + 
                        String.Format(Properties.Resources.DO_YOU_WANT_TO_OPEN),
                    Application.ProductName,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.Yes)
            {
                openCommon(localpath);
            }
            else
            {
                openInExplorer(localpath, program);
            }
        }

        private static void openCommon(string localpath)
        {
            try
            {
                System.Diagnostics.Process.Start(localpath);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName);
            }
        }

        // https://stackoverflow.com/a/696144
        static void openInExplorer(string filePath, string program)
        {
            bool isFile = File.Exists(filePath);
            bool isDir = Directory.Exists(filePath);
            if (!isFile && !isDir)
            {
                exitProgram(string.Format(Properties.Resources.FILE_OR_FOLDER_NOT_EXIST, filePath));
                return;
            }

            try
            {
                if (string.IsNullOrEmpty(program))
                {
                    if (isFile)
                    {
                        // combine the arguments together
                        // it doesn't matter if there is a space after ','
                        string argument = "/select, \"" + filePath + "\"";
                        System.Diagnostics.Process.Start("explorer.exe", argument);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start("explorer.exe",
                            "\"" + filePath + "\"");
                    }
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
            catch (Exception ex)
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

            int secondsep = netfile.IndexOf('\\', firstsep + 1);
            if (secondsep <= 0)
                return netfile;

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
            if (server == null)
                exitProgram(Properties.Resources.SERVER_NULL);
            if (server.ToLower() != Environment.MachineName.ToLower())
            {
                return null;
            }

            ShareCollection shi = ShareCollection.LocalShares;
            if (shi == null)
                exitProgram(Properties.Resources.SHARECOLLECTION_NULL);

            string serverandshare = getServerAndShare(netfile);
            if (serverandshare == null)
                exitProgram(Properties.Resources.SERVERANDSHARE_NULL);
            foreach (Share si in shi)
            {
                if (si.ShareType == ShareType.Disk && si.IsFileSystem)
                {
                    if (serverandshare.ToLower() == si.ToString().ToLower())
                    {
                        return ReplaceFirst(netfile, serverandshare, si.Path);
                    }
                }
            }
            return null;
        }
    }
}
