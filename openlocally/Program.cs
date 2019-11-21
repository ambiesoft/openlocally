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
        static void Main()
        {
            string localpath = TestShares(@"\\thexp\Share\chromium\videos\023\Rendering in WebKit.mp4-");
            if (localpath == null)
                exitProgram("aaa");

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormMain());
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
        static string TestShares(string netfile)
        {
            string server = getServer(netfile);
            if (server == null || server.ToLower() != Environment.MachineName.ToLower())
                exitProgram("aaa");
            ShareCollection shi = ShareCollection.LocalShares;
            if (shi == null)
                exitProgram("aaa");

            string serverandshare = getServerAndShare(netfile);
            if (serverandshare == null)
                exitProgram("aaa");
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
