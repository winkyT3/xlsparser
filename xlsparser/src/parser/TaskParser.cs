using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace xlsparser
{
    class TaskParser : BaseXlsParser
    {
        public override void PostProcessTableList(List<Table> table_list)
        {
            // add sceneobject value
            {
                int condition_key_index = 0;
                int c_param1_key_index = 0;
                List<KeyT> key_list = table_list[0].keyList;
                for (int i = 0; i < key_list.Count; ++i)
                {
                    KeyT key_T = key_list[i];
                    if ("accept_npc" == key_T.key || "commit_npc" == key_T.key)
                    {
                        key_T.keyType = KEY_TYPE.NPC_OBJ;
                    }

                    if ("condition" == key_T.key)
                    {
                        condition_key_index = i;
                    }

                    if ("c_param1" == key_T.key)
                    {
                        c_param1_key_index = i;
                    }
                }

                {
                    KeyT key_T = new KeyT();
                    key_T.key = "target_obj";
                    key_T.keyType = KEY_TYPE.SCENE_OBJ_LIST;
                    key_T.outFlag = ConfigIni.OUTPUT_C.ToString();
                    key_list.Add(key_T);
                }

                List<List<object>> item_list = table_list[0].itemList;
                for (int i = 0; i < item_list.Count; ++i)
                {
                    int condition = Convert.ToInt32(item_list[i][condition_key_index]);
                    int c_param1 = Convert.ToInt32(item_list[i][c_param1_key_index]);
                    item_list[i].Add(string.Format("{0},{1}", condition, c_param1));
                }
            }

            base.PostProcessTableList(table_list);
        }

        public override string ConvertToClientLua(List<Table> table_list)
        {
            if (!SceneObjects.Instance.ReadAllObjects())
            {
                return "";
            }

            return base.ConvertToClientLua(table_list);
        }

        public override XDocument ConvertToServerXml(List<Table> table_list)
        {
            {
                Table table = table_list[0];
                List<List<object>> item_list = table.itemList;
                for (int i = 0; i < item_list.Count; ++ i)
                {
                    List<object> val_list = item_list[i];
                    for (int j = 0; j < val_list.Count; j++)
                    {
                        if (string.IsNullOrEmpty(val_list[j].ToString()))
                        {
                            val_list[j] = 0;
                        }
                    }
                }
            }

            XDocument xmldoc = new XDocument();
            XElement root_node = new XElement("config");
            xmldoc.Add(root_node);

            foreach (Table table in table_list)
            {
                XElement node = XmlBuilder.ConvertTableToXmlElement(table);
                if (null != node)
                {
                    root_node.Add(node);
                }
            }

            return xmldoc;
        }
    }
}
