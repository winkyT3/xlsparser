using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace xlsparser
{
    class BossSkillConditionParser : BaseXlsParser
    {
        protected override bool ParseSheetHeader(ISheet sheet)
        {
            this.header.sheetStartIndex = 0;
            this.header.startLine = 10;
            this.header.tableNameList.Add("Sheet1");

            return true;
        }

        protected override bool ParseSheetData(ISheet sheet, Table ret_table)
        {
            if (null == sheet)
            {
                return false;
            }

            int col_num = 1000;
            int row_num = 100000;
            int start_row_index = this.header.startLine - 1;

            // key
            {
                IRow key_row = sheet.GetRow(start_row_index);
                for (int i = 0; i < col_num; ++i)
                {
                    ICell key_col = key_row.GetCell(i);
                    if (null == key_col)
                    {
                        break;
                    }

                    KeyT key_T = new KeyT();
                    key_T.outFlag = "s";
                    key_T.key = key_col.ToString();
                    key_T.keyType = KEY_TYPE.NORMA;

                    ret_table.keyList.Add(key_T);
                }
            }

            List<Object> value_list = null;
            int condition_id = -1;
            int real_col_num = ret_table.keyList.Count;

            for (int i = start_row_index + 1; i < row_num; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (null == row)
                {
                    continue; // read end
                }

                ICell cell = row.GetCell(0);
                if (null != cell && cell.ToString().Equals("END"))
                {
                    break;
                }

                cell = row.GetCell(0);
                if (null != cell && !string.IsNullOrEmpty(cell.ToString()) && 
                    Convert.ToInt32(cell.NumericCellValue) != condition_id)
                {
                    condition_id = Convert.ToInt32(cell.NumericCellValue);
                    value_list = new List<object>();
                    ret_table.itemList.Add(value_list);
                }

                for (int j = 0; j < real_col_num; ++j)
                {
                    KeyT key_T = ret_table.keyList[j];

                    cell = row.GetCell(j);
                    if (null != cell && !string.IsNullOrEmpty(cell.ToString()))
                    {
                        object value = null;
                        if (CellType.Numeric == cell.CellType)
                        {
                            value = cell.NumericCellValue;
                        }
                        else if (CellType.String == cell.CellType)
                        {
                            value = cell.StringCellValue;
                        }
                        else if (CellType.Formula == cell.CellType)
                        {
                            ICell new_cell = formulaEvaluator.EvaluateInCell(cell);
                            if (CellType.Numeric == new_cell.CellType)
                            {
                                value = new_cell.NumericCellValue;
                            }
                            else if (CellType.String == new_cell.CellType)
                            {
                                value = new_cell.ToString();
                            }
                        }

                        if (value != null)
                        {
                            if (j < value_list.Count && null != value_list[j] &&
                                (key_T.key.Equals("cond") || key_T.key.Equals("skill_id")))
                            {
                                value_list[j] = value_list[j].ToString() + "|" + value.ToString();
                            }
                            else
                            {
                                value_list.Add(value);
                            }
                        }
                        else
                        {
                            value_list.Add(string.Empty);
                        }
                    }
                }
            }

            return true;
        }

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

                string cond_val = "";
                string skill_id_val = "";

                for (int i = 0; i < table.keyList.Count; ++i)
                {
                    KeyT key_T = table.keyList[i];
                    if (key_T.key.Equals("cond"))
                    {
                        cond_val = val_list[i].ToString();
                    }
                    else if (key_T.key.Equals("skill_id"))
                    {
                        skill_id_val = val_list[i].ToString();
                    }
                    else if (key_T.key.Equals("comment"))
                    {
                        continue;
                    }
                    else
                    {
                        XmlBuilder.SetValueInNode(root_node, key_T, val_list[i]);
                    }
                }

                XElement cond_list_node = new XElement("cond_list");
                root_node.Add(cond_list_node);

                string[] conds = cond_val.Split('|');
                string[] skill_ids = skill_id_val.Split('|');

                if (conds.Length != skill_ids.Length)
                {
                    return false;
                }

                for (int i = 0; i < conds.Length; i++)
                {
                    XElement cond_node = new XElement("cond");

                    {
                        cond_list_node.Add(cond_node);
                        string[] ary = conds[i].Split('#');
                        cond_node.SetElementValue("cond_type", ary[0]);
                        for (int j = 1; j < 5; ++ j)
                        {
                            int param = j < ary.Length ? Convert.ToInt32(ary[j]) : 0;
                            cond_node.SetElementValue(string.Format("param{0}", j - 1), param);
                        }
                    }

                    // skill_id
                    {
                        XElement skill_list_node = new XElement("skill_list");
                        cond_node.Add(skill_list_node);

                        string[] ary = skill_ids[i].Split('#');
                        for (int j = 0; j < ary.Length; ++ j)
                        {
                            XElement skill_node = new XElement("skill");
                            skill_list_node.Add(skill_node);
                            skill_node.SetElementValue("skill_id", ary[j]);
                        }
                    }
                }

                string path = string.Format("{0}/gameworld/bossskillcondition/{1}.xml", ConfigIni.XmlDir, val_list[0]);

                Writer.Instance.WriteXml(path, doc, false);
            }

            Command.Instance.AddSvnAddFilePath(string.Format("{0}/gameworld/bossskillcondition/", ConfigIni.XmlDir));
            Command.Instance.AddSvnCommitFilePath(string.Format("{0}/gameworld/bossskillcondition/", ConfigIni.XmlDir));

            // bossskillconditionmanager
            {
                XDocument doc = new XDocument();
                XElement root_node = new XElement("bossskillconditionmanager");
                doc.Add(root_node);

                foreach (List<object> val_list in table.itemList)
                {
                    XElement path_node = new XElement("path");
                    path_node.SetValue(string.Format("bossskillcondition/{0}.xml", val_list[0]));
                    root_node.Add(path_node);
                }

                string path = string.Format("{0}/gameworld/bossskillconditionmanager.xml", ConfigIni.XmlDir);
                Command.Instance.AddSvnAddFilePath(path);

                Writer.Instance.WriteXml(path, doc);
                Command.Instance.AddSvnCommitFilePath(string.Format("{0}/gameworld/bossskillcondition/", ConfigIni.XmlDir));
            }

            return true;
        }
    }
}
