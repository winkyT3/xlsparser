using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace xlsparser
{
    class XmlBuilder
    {
        public static XElement ConvertTableToXmlElement(Table table)
        {
            if (!string.IsNullOrEmpty(table.outFlag)
                && table.outFlag.IndexOf(ConfigIni.OUTPUT_S) < 0)
            {
                return null;
            }

            XElement root_node = new XElement(table.name);

            foreach (List<object> val_list in table.itemList)
            {
                XElement node = new XElement("data");
                for (int i = 0; i < table.keyList.Count; ++i)
                {
                    SetValueInNode(node, table.keyList[i], val_list[i]);
                }

                root_node.Add(node);
            }

            return root_node;
        }

        public static void SetValueInNode(XElement node, KeyT key_T, object value)
        {
            if (key_T.outFlag.IndexOf(ConfigIni.OUTPUT_S) < 0)
            {
                return;
            }

            if (KEY_TYPE.ITEM == key_T.keyType)
            {
                node.Add(GetItemNode(key_T.key, value.ToString()));
            }
            else if (KEY_TYPE.ITEM_LIST == key_T.keyType)
            {
                node.Add(GetItemListNode(key_T.key, value.ToString()));
            }
            else if (KEY_TYPE.DROP_ID == key_T.keyType)
            {
                node.Add(GetDropIdListNode(key_T.key, value.ToString()));
            }
            else if (KEY_TYPE.EFFECT == key_T.keyType)
            {
                node.Add(GetEffectListNode(key_T.key, value.ToString()));
            }
            else
            {
                node.SetElementValue(key_T.key, value);
            }
        }

        private static XElement GetItemNode(string key, string val)
        {
            string[] key_ary = key.Split(',');
            
            XElement node = new XElement(key_ary[0]);

            if (string.IsNullOrEmpty(val) || val.Equals("0"))
            {
                node.SetElementValue("item_id", 0);
                node.SetElementValue("num", "");
                node.SetElementValue("is_bind", "");
            }
            else
            {
                string[] val_ary = val.Split(':');
                node.SetElementValue("item_id", val_ary[0]);
                node.SetElementValue("num", val_ary[1]);
                node.SetElementValue("is_bind", val_ary[2]);
            }

            return node;
        }

        private static XElement GetItemListNode(string key, string val)
        {
            string[] key_ary = key.Split(',');
            string[] list_ary = val.Split(',');

            XElement list_node = new XElement(string.Format("{0}_list", key_ary[0]));

            for (int i = 0; i < list_ary.Length; ++i)
            {
                string[] val_ary = list_ary[i].Split(':');

                if (3 == val_ary.Length)
                {
                    XElement node = new XElement(key_ary[0]);
                    node.SetElementValue("item_id", val_ary[0]);
                    node.SetElementValue("num", val_ary[1]);
                    node.SetElementValue("is_bind", val_ary[2]);
                    list_node.Add(node);
                }
            }

            return list_node;
        }

        private static XElement GetDropIdListNode(string key, string val)
        {
            XElement list_node = new XElement(key);
            string[] id_list = val.Split('|');

            for (int i = 0; i < id_list.Length; ++i)
            {
                XElement node = new XElement("dropid");
                node.SetValue(id_list[i]);
                list_node.Add(node);
            }

            return list_node;
        }

        private static XElement GetEffectListNode(string key, string val)
        {
            XElement list_node = new XElement("effect_list");
            XElement effect_node = new XElement("effect");
            list_node.Add(effect_node);

            string[] param_list = val.Split('#');
            effect_node.SetElementValue("effect_type", param_list[0]);

            for (int i = 1; i < 7; ++i)
            {
                effect_node.SetElementValue("param" + (i - 1), i < param_list.Length ? Convert.ToInt32(param_list[i]) : 0);
            }

            return list_node;
        }
    }
}
