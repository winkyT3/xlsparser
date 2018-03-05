using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace xlsparser
{
    class DropParser : BaseXlsParser
    {
        public override bool Parse(List<ISheet> sheet_list, List<Table> table_list)
        {
            this.header = new SheetHeader();
            table_list.Clear();

            this.header.sheetStartIndex = 0;
            this.header.fileName = "drop";
            this.header.serverPath = "gameworld/drop";
            this.header.startLine = 10;
            this.header.tableNameList.Add("sheet1");

            Table table = new Table();
            table.name = "drop";
            table_list.Add(table);

            if (!this.ParseSheetData(sheet_list[0], table))
            {
                return false;
            }

            return true;
        }

        protected override bool ParseSheetData(ISheet sheet, Table ret_table)
        {
            if (null == sheet)
            {
                return false;
            }

            int col_num = 1000;
            int row_num = 10000000;

            // list value
            for (int i = this.header.startLine - 1; i < row_num; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (null == row)
                {
                    break; // read end
                }

                List<object> value_list = new List<object>();

                bool is_end = false;
                for (int j = 0; j < col_num; ++j)
                {
                    ICell cell = row.GetCell(j);
                    if (null == cell || string.IsNullOrEmpty(cell.ToString()))
                    {
                        if (0 == j)
                        {
                            is_end = true;
                        }

                        break;
                    }
                    else
                    {
                        if (CellType.Numeric == cell.CellType)
                        {
                            value_list.Add(cell.NumericCellValue);
                        }
                        else if (CellType.String == cell.CellType)
                        {
                            value_list.Add(cell.StringCellValue);
                        }
                        else if (CellType.Formula == cell.CellType)
                        {
                            ICell new_cell = formulaEvaluator.EvaluateInCell(cell);
                            if (CellType.Numeric == new_cell.CellType)
                            {
                                value_list.Add(new_cell.NumericCellValue);
                            }
                            else if (CellType.String == new_cell.CellType)
                            {
                                value_list.Add(new_cell.ToString());
                            }
                        }
                    }
                }

                if (is_end)
                {
                    break; // read end
                }

                ret_table.itemList.Add(value_list);
            }

            return true;
        }

        public override bool BuildServerXml(List<Table> table_list)
        {
            if (table_list.Count < 1)
            {
                return false;
            }

            Table table = table_list[0];

            foreach (List<object> val_list in table.itemList)
            {
                XDocument doc = new XDocument();
                XElement root_node = new XElement("drop");
                doc.Add(root_node);

                root_node.SetElementValue("drop_id", val_list[0]);

                XElement prop_list_node = new XElement("drop_item_prob_list");
                root_node.Add(prop_list_node);

                for (int i = 2; i < val_list.Count; ++i)
                {
                    string[] ary = val_list[i].ToString().Split('#');
                    if (5 != ary.Length)
                    {
                        return false;
                    }

                    XElement drop_item_prob_node = new XElement("drop_item_prob");
                    drop_item_prob_node.SetElementValue("item_id", ary[0]);
                    drop_item_prob_node.SetElementValue("is_bind", ary[1]);
                    drop_item_prob_node.SetElementValue("prob", ary[2]);
                    drop_item_prob_node.SetElementValue("num", ary[3]);
                    drop_item_prob_node.SetElementValue("broadcast", ary[4]);

                    prop_list_node.Add(drop_item_prob_node);
                }

                string path = string.Format("{0}/{1}/{2}.xml", ConfigIni.XmlDir, header.serverPath, val_list[0]);
                Writer.Instance.WriteXml(path, doc, false);
            }

            Command.Instance.AddSvnAddFilePath(string.Format("{0}/{1}", ConfigIni.XmlDir, header.serverPath));
            Command.Instance.AddSvnCommitFilePath(string.Format("{0}/{1}", ConfigIni.XmlDir, header.serverPath));

            // drop manager
            {
                XDocument doc = new XDocument();
                XElement root_node = new XElement("dropmanager");
                doc.Add(root_node);

                foreach (List<object> val_list in table.itemList)
                {
                    XElement path_node = new XElement("path");
                    path_node.SetValue(string.Format("drop/{0}.xml", val_list[0]));
                    root_node.Add(path_node);
                }

                string path = string.Format("{0}/gameworld/dropmanager.xml", ConfigIni.XmlDir);
                Command.Instance.AddSvnAddFilePath(path);

                Writer.Instance.WriteXml(path, doc);
                Command.Instance.AddSvnCommitFilePath(path);
            }

            return true;
        }
    }
}
