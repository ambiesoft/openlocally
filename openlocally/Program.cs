using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Trinet.Networking;

namespace openlocally
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            if (args.Length == 0)
                exitProgram(Properties.Resources.NO_ARGUMENTS);

            string inputpath = args[0];
            string fullpath = Path.GetFullPath(inputpath);
            string localpath;
            if (Char.IsLetter(fullpath[0]) &&
                fullpath[1] == ':' &&
                fullpath[2] == '\\')
            {
                localpath = fullpath;
            }
            else
            {
                localpath = GetLocalPath(inputpath);
                if (localpath == null)
                    exitProgram(Properties.Resources.LOCAL_PATH_NOT_FOUND);
            }
                
            openInExplorer(localpath);
            
            
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new FormMain());
        }

        // https://stackoverflow.com/a/696144
        static void openInExplorer(string filePath)
        {
            if (!File.Exists(filePath) && !Directory.Exists(filePath))
            {
                exitProgram(string.Format(Properties.Resources.FILE_OR_FOLDER_NOT_EXIST, filePath));
                return;
            }

            // combine the arguments together
            // it doesn't matter if there is a space after ','
            string argument = "/select, \"" + filePath + "\"";

            System.Diagnostics.Process.Start("explorer.exe", argument);
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
