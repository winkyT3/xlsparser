using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Xml;
using System.Xml.Linq;
using System.Drawing;
using System.Text.RegularExpressions;

namespace xlsparser
{
    enum KEY_TYPE
    {
        NORMA = 0,
        MAIN_KEY,
        ITEM,
        ITEM_LIST,
        NPC_OBJ,
        SCENE_OBJ_LIST,
        DROP_ID,
        EFFECT,
        SPLIT_LIST,
        STEP_LIST,
    }

    enum TABLE_TYPE
    {
        NORMAL = 0,
        LIST = 1,
        SIMPLE = 2,
    }

    class SheetHeader
    {
        public string fileName;
        public string serverPath;
        public int startLine;
        public List<string> tableNameList = new List<string>();
        public int sheetStartIndex;
    }

    class KeyT
    {
        public string key;
        public KEY_TYPE keyType;
        public string outFlag;

        public KeyT Clone()
        {
            return this.MemberwiseClone() as KeyT;
        }
    }

    class Table
    {
        public string name;
        public string outFlag;
        public TABLE_TYPE tableType;
        public List<KeyT> keyList = new List<KeyT>();
        public List<List<object>> itemList = new List<List<object>>();
    }

    class BaseXlsParser
    {
        protected SheetHeader header;
        public static HSSFFormulaEvaluator formulaEvaluator;

        public BaseXlsParser()
        {

        }

        public string GetFileName()
        {
            return null != this.header ? this.header.fileName : string.Empty;
        }

        public virtual bool Parse(List<ISheet> sheet_list, List<Table> table_list)
        {
            this.header = new SheetHeader();
            table_list.Clear();

            if (sheet_list.Count() <= 1)
            {
                return false;
            }

            if (!this.ParseSheetHeader(sheet_list[0]))
            {
                return false;
            }

            for (int i = 0; i < this.header.tableNameList.Count; ++i)
            {
                Table table = new Table();
                string[] temp_ary = this.header.tableNameList[i].Split(',');

                if (temp_ary.Length >= 2)
                {
                    table.outFlag = temp_ary[1];
                }

                table.name = temp_ary[0].Trim(' ');
                table_list.Add(table);

                if (!this.ParseSheetData(sheet_list[this.header.sheetStartIndex + i], table))
                {
                    return false;
                }
            }

            return true;
        }

        protected virtual bool ParseSheetHeader(ISheet sheet)
        {
            if (sheet.PhysicalNumberOfRows < 2)
            {
                return false;
            }

            this.header.sheetStartIndex = 1;
            IRow row = sheet.GetRow(1);

            for (int i = 0; i < row.LastCellNum; ++i)
            {
                ICell cell = row.GetCell(i);
                if (null == cell)
                {
                    break;
                }

                if (0 == i)
                {
                    this.header.fileName = cell.ToString();
                }
                else if (1 == i)
                {
                    this.header.serverPath = cell.ToString();
                }
                else if (2 == i)
                {
                    this.header.startLine = Convert.ToInt32(cell.ToString());
                }
                else if (!string.IsNullOrEmpty(cell.ToString()))
                {
                    this.header.tableNameList.Add(cell.ToString());
                }
            }

            return true;
        }

        protected virtual bool ParseSheetData(ISheet sheet, Table ret_table)
        {
            if (null == sheet)
            {
                return false;
            }

            int row_num = sheet.LastRowNum;
            int start_row_index = this.header.startLine - 1;
            if (start_row_index < 0 || 
                (row_num > 0 && start_row_index >= row_num))
            {
                return false;
            }

            // LastRowNum have bug
            int col_num = 1000;

            // key
            {
                IRow out_flag_row = sheet.GetRow(start_row_index);
                IRow key_row = sheet.GetRow(start_row_index + 2);
                for (int i = 0; i < col_num; ++i)
                {
                    if (null == out_flag_row)
                    {
                        break;
                    }

                    ICell flag_col = out_flag_row.GetCell(i);
                    ICell key_col = key_row.GetCell(i);
                    if (null == flag_col || null == key_col)
                    {
                        break;
                    }

                    string out_flag = flag_col.ToString();
                    string key = key_col.ToString();
                    key = key.Trim(' ');
                    key = key.Trim('\n');

                    if (string.IsNullOrEmpty(out_flag) && string.IsNullOrEmpty(key))
                    {
                        break;
                    }

                    if (!string.IsNullOrEmpty(out_flag) && !string.IsNullOrEmpty(key))
                    {
                        KeyT key_T = new KeyT();
                        key_T.outFlag = out_flag;
                        key_T.key = key;

                        string[] flags = out_flag.IndexOf("，") >= 0 ? out_flag.Split('，') : out_flag.Split(',');
                        key_T.outFlag = flags[0];

                        if (flags.Length > 1)
                        {
                            if (flags[1].Equals(ConfigIni.INDEX_FLAG) && 0 == i)
                            {
                                key_T.keyType = KEY_TYPE.MAIN_KEY;
                            }
                            else if (flags[1].Equals(ConfigIni.ITEMLIST_FLAG))
                            {
                                key_T.keyType = KEY_TYPE.ITEM_LIST;
                            }
                            else if (flags[1].Equals(ConfigIni.ITEM_FLAG))
                            {
                                key_T.keyType = KEY_TYPE.ITEM;
                            }
                            else if (flags[1].Equals(ConfigIni.DROP_ID_FLAG))
                            {
                                key_T.keyType = KEY_TYPE.DROP_ID;
                            }
                            else if (flags[1].Equals(ConfigIni.EFFECT_FLAG))
                            {
                                key_T.keyType = KEY_TYPE.EFFECT;
                            }
                            else if (flags[1].Equals(ConfigIni.SPLITLIST_FLAG))
                            {
                                key_T.keyType = KEY_TYPE.SPLIT_LIST;
                            }
                            else if (flags[1].Equals(ConfigIni.STEPLIST_FLAG))
                            {
                                key_T.keyType = KEY_TYPE.STEP_LIST;
                            }
                            else
                            {
                                key_T.keyType = KEY_TYPE.NORMA;
                            }
                        }

                        ret_table.keyList.Add(key_T);
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            // real col num
            col_num = ret_table.keyList.Count;
            // LastRowNum have bug
            row_num = 10000000;
            
            // list value
            for (int i = start_row_index + 3; i < row_num; ++i)
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
                        value_list.Add(string.Empty);
                        if (0 == j)
                        {
                            is_end = true;
                        }
                    }
                    else
                    {
                        if (CellType.Numeric == cell.CellType)
                        {
                            value_list.Add(cell.NumericCellValue);
                        }
                        else if (CellType.String == cell.CellType)
                        {
                            if (Regex.Match(cell.StringCellValue, "^[-+]?[0-9]*\\.?[0-9]+$").Success)
                            {
                                value_list.Add(Convert.ToDouble(cell.StringCellValue));
                            }
                            else
                            {
                                value_list.Add(cell.StringCellValue);
                            }
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
                                if (Regex.Match(cell.StringCellValue, "^[-+]?[0-9]*\\.?[0-9]+$").Success)
                                {
                                    value_list.Add(Convert.ToDouble(cell.StringCellValue));
                                }
                                else
                                {
                                    value_list.Add(cell.StringCellValue);
                                }
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

        public virtual void PostProcessTableList(List<Table> table_list)
        {
            LuaBuilder.SimplifiedTableList(table_list);
        }

        public virtual bool BuildClientLua(List<Table> table_list)
        {
            string lua_str = this.ConvertToClientLua(table_list);
            if (string.IsNullOrEmpty(lua_str))
            {
                return false;
            }

            string path = string.Format("{0}/{1}_auto.lua", ConfigIni.LuaDir, header.fileName);

            Command.Instance.AddSvnAddFilePath(path);
            Writer.Instance.WriteFile(path, lua_str);
            
            Command.Instance.AddSvnCommitFilePath(path);

            return true;
        }

        public virtual string ConvertToClientLua(List<Table> table_list)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("return {\n");
            for (int i = 0; i < table_list.Count; ++ i)
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

        public virtual bool BuildServerXml(List<Table> table_list)
        {
            XDocument xmldoc = this.ConvertToServerXml(table_list);

            string path = string.Format("{0}/{1}/{2}.xml", ConfigIni.XmlDir, header.serverPath, header.fileName);
            Command.Instance.AddSvnAddFilePath(path);

            Writer.Instance.WriteXml(path, xmldoc);
            Command.Instance.AddSvnCommitFilePath(path);

            return true;
        }

        public virtual XDocument ConvertToServerXml(List<Table> table_list)
        {
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
