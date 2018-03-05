using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace xlsparser
{
    class MonsterParser : BaseXlsParser
    {
        public override bool BuildServerXml(List<Table> table_list)
        {
            if (table_list.Count <= 0)
            {
                return false;
            }

            Table table = table_list[0];
            foreach (List<object> val_list in table.itemList)
            {
                XDocument doc = new XDocument();
                XElement root_node = new XElement("config");
                doc.Add(root_node);

                for (int i = 0; i < table.keyList.Count; ++i)
                {
                    XmlBuilder.SetValueInNode(root_node, table.keyList[i], val_list[i]);
                }

                string path = string.Format("{0}/{1}/{2}.xml", ConfigIni.XmlDir, header.serverPath, val_list[0]);
                Writer.Instance.WriteXml(path, doc, false);
            }

            Command.Instance.AddSvnAddFilePath(string.Format("{0}/{1}/", ConfigIni.XmlDir, header.serverPath));
            Command.Instance.AddSvnCommitFilePath(string.Format("{0}/{1}/", ConfigIni.XmlDir, header.serverPath));

            // monster manager
            {
                XDocument doc = new XDocument();
                XElement root_node = new XElement("config");
                doc.Add(root_node);

                XElement monster_list_node = new XElement("monster_list");
                root_node.Add(monster_list_node);

                foreach (List<object> val_list in table.itemList)
                {
                    XElement path_node = new XElement("path");
                    path_node.SetValue(string.Format("monster/{0}.xml", val_list[0]));
                    monster_list_node.Add(path_node);
                }

                string path = string.Format("{0}/gameworld/monstermanager.xml", ConfigIni.XmlDir);
                Command.Instance.AddSvnAddFilePath(path);
                Writer.Instance.WriteXml(path, doc);
                Command.Instance.AddSvnCommitFilePath(path);
            }

            return true;
        }
    }
}
