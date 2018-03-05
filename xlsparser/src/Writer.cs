using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml;
using System.Drawing;

namespace xlsparser
{
    class Writer : Singleton<Writer>
    {
        public void WriteFile(string path, string conent, bool is_log = true)
        {
            if (is_log)
            {
                Command.Instance.PrintLog("create lua : " + path, Color.Green);
            }

            StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
            using (sw)
            {
                sw.Write(conent);
                sw.WriteLine("\n");
                sw.Flush();
                sw.Close();
            }
        }

        public void WriteXml(string path, XDocument doc, bool is_log = true)
        {
            if (is_log)
            {
                Command.Instance.PrintLog("create xml : " + path, Color.Green);
            }

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.NewLineChars = "\n";
            settings.Indent = true;
            settings.Encoding = new UTF8Encoding(false);

            StreamWriter tw = new StreamWriter(path, false, new UTF8Encoding(false));
            tw.NewLine = "\n";

            XmlWriter xw = XmlWriter.Create(tw, settings);
            using (tw)
            {
                using (xw)
                {
                    doc.Save(xw);
                }

                tw.WriteLine();
            }
        }

        public void WriteRsa(string path)
        {
            StreamWriter sw = new StreamWriter(path, false);
            using (sw)
            {
                sw.Write(ConfigIni.GetRsaKey());
                sw.Flush();
                sw.Close();
            }
        }

        public void DelRsa(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
}
