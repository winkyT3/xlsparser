using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Xml.Linq;

namespace xlsparser
{
    class Builder : Singleton<Builder>
    {
        private NormalXlsParser moduleParser = new NormalXlsParser();
        private ItemParser itemParser = new ItemParser();
        private MonsterSkillXlsParser monsterSkillXlsParser = new MonsterSkillXlsParser();
        private TaskParser taskParser = new TaskParser();
        private DropParser dropParser = new DropParser();
        private MonsterParser monsterParser = new MonsterParser();
        private BossSkillConditionParser bossSkillConditionParser = new BossSkillConditionParser();

        public bool BuildClient(XLS_PARSER_TYPE parser_type, List<ISheet> sheet_list)
        {
            BaseXlsParser parser = this.GetParser(parser_type);
            if (null == parser)
            {
                return false;
            }

            List<Table> table_list = new List<Table>();
            if (!parser.Parse(sheet_list, table_list))
            {
                return false;
            }

            parser.PostProcessTableList(table_list);

            return parser.BuildClientLua(table_list);
        }

        public bool BuildServer(XLS_PARSER_TYPE parser_type, List<ISheet> sheet_list)
        {
            BaseXlsParser parser = this.GetParser(parser_type);
            if (null == parser)
            {
                return false;
            }

            List<Table> table_list = new List<Table>();
            if (!parser.Parse(sheet_list, table_list))
            {
                return false;
            }

            return parser.BuildServerXml(table_list);
        }

        private BaseXlsParser GetParser(XLS_PARSER_TYPE parser_type)
        {
            if (XLS_PARSER_TYPE.NORMAL == parser_type)
            {
                return this.moduleParser;
            }
            else if (XLS_PARSER_TYPE.ITEM == parser_type)
            {
                return this.itemParser;
            }
            else if (XLS_PARSER_TYPE.MONSTER_SKILL == parser_type)
            {
                return this.monsterSkillXlsParser;
            }
            else if (XLS_PARSER_TYPE.TASK == parser_type)
            {
                return this.taskParser;
            }
            else if (XLS_PARSER_TYPE.DROP == parser_type)
            {
                return this.dropParser;
            }
            else if (XLS_PARSER_TYPE.MONSTER == parser_type)
            {
                return this.monsterParser;
            }
            else if (XLS_PARSER_TYPE.BOSS_SKILL_CONDITION == parser_type)
            {
                return this.bossSkillConditionParser;
            }

            return null;
        }

        public bool BuildItemIndex(string dir_path)
        {
            BaseXlsParser parser = this.GetParser(XLS_PARSER_TYPE.ITEM);
            if (null == parser)
            {
                return false;
            }

            string[] xls_list = { "W-装备.xls" , "W-被动消耗类.xls", "W-主动使用消耗类.xls", "W-礼包类.xls", "W-虚拟类.xls" };

            XDocument xmldoc = new XDocument();
            XElement root_node = new XElement("config");
            xmldoc.Add(root_node);

            for (int i = 0; i < xls_list.Length; ++ i)
            {
                List<ISheet> sheet_list = new List<ISheet>();
                if (!XlsReader.Instance.ReadExcel(dir_path + xls_list[i], sheet_list))
                {
                    return false;
                }

                List<Table> temp_list = new List<Table>();
                if (!parser.Parse(sheet_list, temp_list))
                {
                    return false;
                }

                string file_name = parser.GetFileName();

                foreach (Table table in temp_list)
                {
                    XElement table_node = new XElement(table.name);
                    root_node.Add(table_node);

                    foreach (var val_list in table.itemList)
                    {
                        XElement path_node = new XElement("path");
                        path_node.SetValue(string.Format("{0}/{1}.xml", file_name, val_list[0], val_list[1]));
                        table_node.Add(path_node);
                    }
                }
            }

            string path = string.Format("{0}/gameworld/itemmanager.xml", ConfigIni.XmlDir);
            Command.Instance.AddSvnAddFilePath(path);

            Writer.Instance.WriteXml(path, xmldoc);
            Command.Instance.AddSvnCommitFilePath(path);

            return true;
        }
    }
}
