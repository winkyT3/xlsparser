using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace xlsparser
{
    public enum XLS_PARSER_TYPE
    {
        NORMAL = 0,
        ITEM,
        ITEM_INDEX,
        MONSTER_SKILL,
        TASK,
        DROP,
        MONSTER,
        BOSS_SKILL_CONDITION,
    }

    public class ConfigIni
    {
        public static char OUTPUT_C = 'c';
        public static char OUTPUT_S = 's';
        public static string INDEX_FLAG = "index";
        public static string ITEMLIST_FLAG = "itemlist";
        public static string ITEM_FLAG = "item";
        public static string DROP_ID_FLAG = "dropid";
        public static string EFFECT_FLAG = "effect";
        public static string SPLITLIST_FLAG = "splitlist";
        public static string STEPLIST_FLAG = "steplist";

        public static int WinWidth = 1200;
        public static int WinHeight = 600;
        public static string ProjectName = "";
        public static string Database = "";
        public static string XlsDir = "";
        public static string LuaDir = "";
        public static string XmlDir = "";

        [DllImport("kernel32")]//返回0表示失败，非0为成功
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public static bool ReadIniConifg()
        {
            string path = "./config.ini";
            if (!File.Exists(path))
            {
                MessageBox.Show("config.ini不存在");
                return false;
            }

            const int MAX_SIZE = 1024;
            StringBuilder sb = new StringBuilder(MAX_SIZE);

            string content = File.ReadAllText(path);
            StreamWriter sw = new StreamWriter(path, false, new UTF8Encoding(false));
            using (sw)
            {
                sw.Write(content);
                sw.Flush();
                sw.Close();
            }

            //sb.Clear();
            //GetPrivateProfileString("Default", "win_width", "", sb, MAX_SIZE, path);
            //ConfigIni.WinWidth = Math.Max(Convert.ToInt32(sb.ToString()), 100);

            sb.Clear();
            GetPrivateProfileString("Default", "win_height", "", sb, MAX_SIZE, path);
            ConfigIni.WinHeight = Math.Max(Convert.ToInt32(sb.ToString()), 100);


            sb.Clear();
            GetPrivateProfileString("Default", "project_name", "", sb, MAX_SIZE, path);
            ConfigIni.ProjectName = sb.ToString();

            sb.Clear();
            GetPrivateProfileString("Default", "database", "", sb, MAX_SIZE, path);
            ConfigIni.Database = sb.ToString();
            
            sb.Clear();
            GetPrivateProfileString("Default", "xls_dir", "", sb, MAX_SIZE, path);
            ConfigIni.XlsDir = sb.ToString();

            sb.Clear();
            GetPrivateProfileString("Default", "lua_dir", "", sb, MAX_SIZE, path);
            ConfigIni.LuaDir = sb.ToString();

            sb.Clear();
            GetPrivateProfileString("Default", "xml_dir", "", sb, MAX_SIZE, path);
            ConfigIni.XmlDir = sb.ToString();

            if (string.IsNullOrEmpty(ConfigIni.ProjectName))
            {
                MessageBox.Show("需指定项目名字，请先配置config.ini");
                return false;
            }

            if (!Directory.Exists(ConfigIni.XlsDir)
                || !Directory.Exists(ConfigIni.LuaDir)
                || !Directory.Exists(ConfigIni.XmlDir))
            {
                MessageBox.Show("找不到指定的文件夹路径，请先配置config.ini");
                return false;
            }

            return true;
        }


        public static string GetRsaKeySavePath()
        {
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            path = path.Replace("\\", "/") + "/id_rsa.rsa";

            return path;
        }

        public static string GetRsaKey()
        {
            return "-----BEGIN RSA PRIVATE KEY-----\n" +
            "MIIEoQIBAAKCAQEAx9VxgZjD/LEty2GV7UnWnJfM7qdu9Yx1gvpBhopRkssfnykL\n" +
            "FUxEeyDu6ncOivJE51sK1Q2OJZHtwVJHdf+rd0dfetvJS1WRJbux10yr07IAIxT6\n" +
            "8heDDhqDbP1zqYGellE5/WNyjV4rGOb+peA8i9Ov4nm+r0nC9cvjeI8CCQ2/355l\n" +
            "oZT7irIwkvH8et3+WbNeIwctYh1EvEbuZr9wH3TySm8kX/SwjltP90YKFICo99WE\n" +
            "6HHuqn+x9OylYgiVfOm5DWs/L+zcomOOhPv/A7l6IXTDx0HSfmkm4XOC33OyGWG9\n" +
            "sU9fGI0tIi2uRZh1Ch6XhiZIXrqEWDjehrUVOQIBIwKCAQEApZOYj/ObEzO4Os0z\n" +
            "FRFLa848FjL1iaBEHBEu+nKbXF8o1FU1EaWX07wzrFQEvEUUhS4ti/VJ5J172rHj\n" +
            "cGYgW4RHzDJzljD0m5uEubSOXvKSZjX0f3KRKPFllNIAv6XpzQF5MQlBqFVW6L9l\n" +
            "R5yJ8DMOE2wwSBiLmHW8edWT+N4JsL/fgBw5K/QwgIFE2VczVLqMuaIrfw/jcxJ/\n" +
            "gAwt5dh2rb8bVzD1DwgrINPGctUzkr6yMzb8lpo84ts/zjPgrOP6CYzsHciOOE65\n" +
            "S//Elw9z5ZHwB7mMOOOLu1tKgpPQ6q0DnIZnw5iEg1U4PlkJLAr+g67G1CR0Go/r\n" +
            "xWr3pwKBgQD1L1uWfnAkHGZcEtjDMv1XHEi4iSpjEV2EEgJPq/AXTynFcXewBGsy\n" +
            "YEZNNZwc0k3mkD+aH9FPBPi5elzSKRnRvbf060+vIK2SqGRL8Fs1OFWdhK7FHBLi\n" +
            "cvDsOVcX8kAWwt9g99XUq6zMPdnm+QVPtzpD85z/yCekEUDYtD6bqwKBgQDQpfoi\n" +
            "f7NUoy3yPUg+dokRA0u1I8DKyQ2uKDcfVTw75wgiB0XGHnG3hd8Dq7Mp476wTIx0\n" +
            "0W/qYoYzaMUBzU8M3ODUKDyEvyntx7HnXDcjJrskmVx736dGz6Nl4HlF91UR9Ki5\n" +
            "eCbGknfA9l+vjpGeqvpeMWbl3SkWW+Bk219OqwKBgFQQPKigYPZwIxhA+dyGgr7H\n" +
            "3msntsLhYeQjboG3SwCtbWhEKQkmFh/mfoDfLjXEcnrvoMcg2gx2u61OhjltLWx7\n" +
            "j4csG1H8k0g5u/zHb7p5tvQtfb/sXj8C1kJczWdLvjOwhxnuknTNJU1XCOF58zFG\n" +
            "Ipr0cFeyVrvoqITQFXczAoGAa04LnLaz/6Rgxbkd1vsh634JrZ1b1gD/uKb3xvii\n" +
            "onbQ+462K2AdOc/Rx1+d+Ek9fz1PjIj3ul6OKRFPYAMv9/yQ4iNSUuX4TmazQ8kG\n" +
            "adlnjypM0f3+QaydLRRboFNQUmCRSXD9/7kKYzzgwK+4mr5U3/wmSlR7h9d6t4bD\n" +
            "TQcCgYBlBohdx4dMuZ/wFntgd2zkkzOZCpGmyoS/I4KXJgvl0aVvXyRuf0LT2fDF\n" +
            "aq0CgJxk6DkoccV2coBwEiUagsHj+PxTifbHXOAZ3Z93lu/81S/f5IvdZ9BMFnSN\n" +
            "91R/V1l+Rf+pjXM7l2n5DvrT0SFWKlv2bXgKyT1eoyA52v7kPA==\n" +
            "-----END RSA PRIVATE KEY-----";
        }
    }
}
