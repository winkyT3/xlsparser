using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using xlsparser.view;
using System.Collections;
using System.Text.RegularExpressions;

namespace xlsparser
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (!ConfigIni.ReadIniConifg())
            {
                return;
            }

            //   Command.Execute("mkdir 5");

            // Command.Execute("svn commit * -m");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new BuildWin());
        }
    }
}
