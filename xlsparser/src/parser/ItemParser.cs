using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Xml.Linq;

namespace xlsparser
{
    class ItemParser : BaseXlsParser
    {
        protected override bool ParseSheetHeader(ISheet sheet)
        {
            if (!base.ParseSheetHeader(sheet))
            {
                return false;
            }

            this.header.fileName = string.Format("item/{0}", this.header.fileName);
            return true;
        }

        public override void PostProcessTableList(List<Table> table_list)
        {
            // combine table
            {
                List<List<object>> item_list = table_list[0].itemList;
                for (int i = 1; i < table_list.Count; i++)
                {
                    List<List<object>> temp_item_list = table_list[i].itemList;
                    for (int j = 0; j < temp_item_list.Count; ++j)
                    {
                        item_list.Add(temp_item_list[j]);
                    }
                }

                while (table_list.Count > 1)
                {
                    table_list.RemoveAt(table_list.Count - 1);
                }

                table_list[0].tableType = TABLE_TYPE.LIST;
                table_list[0].keyList[0].keyType = KEY_TYPE.MAIN_KEY;
            }

            // handle default table
            {
                Table default_table = LuaBuilder.SimplifiedTable(table_list[0]);
                default_table.tableType = TABLE_TYPE.SIMPLE;
                default_table.name = "default_table";

                KeyT key_T = new KeyT();
                key_T.outFlag = "";
                key_T.key = "default_table";
                key_T.keyType = KEY_TYPE.MAIN_KEY;
                default_table.keyList.Add(key_T);

                table_list.Add(default_table);
            }
        }

        public override string ConvertToClientLua(List<Table> table_list)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("return {\n");
            for (int i = 0; i < table_list.Count; ++i)
            {
                string lua_table = LuaBuilder.GetLuaTable(table_list[i]);
                if (!string.IsNullOrEmpty(lua_table))
                {
                    builder.Append(lua_table);
                    builder.Append(i != table_list.Count - 1 ? ",\n\n" : "");
                }
            }
            builder.Append("\n\n}");
            return builder.ToString();
        }

        public override bool BuildServerXml(List<Table> table_list)
        {
            foreach (Table table in table_list)
            {
                foreach (List<object> val_list in table.itemList)
                {
                    XDocument doc = new XDocument();
                    XElement root_node = new XElement("config");
                    doc.Add(root_node);

                    for (int i = 0; i < table.keyList.Count; ++i)
                    {
                        XmlBuilder.SetValueInNode(root_node, table.keyList[i], val_list[i]);
                    }

                    string path = string.Format("{0}/{1}/{2}.xml", ConfigIni.XmlDir, header.serverPath, val_list[0].ToString());
                    Writer.Instance.WriteXml(path, doc, false);
                }
            }

            Command.Instance.AddSvnAddFilePath(string.Format("{0}/{1}/", ConfigIni.XmlDir, header.serverPath));
            Command.Instance.AddSvnCommitFilePath(string.Format("{0}/{1}/", ConfigIni.XmlDir, header.serverPath));

            return true;
        }
    }
}
