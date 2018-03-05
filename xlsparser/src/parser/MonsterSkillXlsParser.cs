using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Xml.Linq;

namespace xlsparser
{
    class MonsterSkillXlsParser : BaseXlsParser
    {
        protected override bool ParseSheetHeader(ISheet sheet)
        {
            this.header.sheetStartIndex = 0;
            this.header.startLine = 8;
            this.header.fileName = "monsterskill";
            this.header.serverPath = "";
            this.header.tableNameList.Add("skill_list");

            return true;
        }

        public override void PostProcessTableList(List<Table> table_list)
        {
            for (int i = 0; i < table_list.Count; i++)
            {
                table_list[i].keyList[0].keyType = KEY_TYPE.MAIN_KEY;
            }

            table_list.Add(LuaBuilder.SimplifiedTable(table_list[0]));
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
                int skill_id = Convert.ToInt32(val_list[0]);
                XDocument doc = new XDocument();

                XElement root_node = new XElement("Skill");
                doc.Add(root_node);

                bool is_self_pos = false;

                for (int i = 0; i < table.keyList.Count; ++i)
                {
                    KeyT key_T = table.keyList[i];

                    // del effect
                    if ((key_T.key.Equals("effect") || key_T.key.Equals("tipType"))
                        && string.IsNullOrEmpty(val_list[i].ToString()))
                    {
                        continue;
                    }

                    // set skillname output
                    if (key_T.key.Equals("skillname"))
                    {
                        key_T.outFlag = "cs";
                    }

                    if (key_T.key.Equals("Distance") && Convert.ToInt32(val_list[i]) <= 1)
                    {
                        is_self_pos = true;
                    }
            
                    XmlBuilder.SetValueInNode(root_node, key_T, val_list[i]);

                    // special add
                    {
                        if (key_T.key.Equals("Distance"))
                        {
                            KeyT temp_key_T = new KeyT();
                            temp_key_T.key = "ProfLimit";
                            temp_key_T.outFlag = "s";
                            XmlBuilder.SetValueInNode(root_node, temp_key_T, 4);
                        }

                        if (key_T.key.Equals("Range"))
                        {
                            KeyT temp_key_T = new KeyT();
                            temp_key_T.key = "IsSelfPos";
                            temp_key_T.outFlag = "s";
                            int is_self_pos_val = is_self_pos && Convert.ToInt32(val_list[i]) > 0 ? 1 : 0;
                            XmlBuilder.SetValueInNode(root_node, temp_key_T, is_self_pos_val);
                        }
                    }
                }

                string group_name = this.GetSkillGroupName(skill_id);
                string path = string.Format("{0}/gameworld/skill/monsterskills/{1}{2}.xml", ConfigIni.XmlDir, group_name, skill_id);
                
                Writer.Instance.WriteXml(path, doc, false);
            }

            Command.Instance.AddSvnAddFilePath(string.Format("{0}/gameworld/skill/monsterskills/", ConfigIni.XmlDir));
            Command.Instance.AddSvnCommitFilePath(string.Format("{0}/gameworld/skill/monsterskills/", ConfigIni.XmlDir));

            // monster manager
            {
                XDocument doc = new XDocument();
                XElement root_node = new XElement("skills");
                doc.Add(root_node);

                Dictionary<string, XElement> group_dic = new Dictionary<string, XElement>();

                // init group
                {
                    string[] group_name_list = { "CommonSkillToEnemy", "RangeCommonSkillToEnemyPos", "CommonSkillToSelf",
                                                "RangeCommonSkillToSelfPos", "FaZhenSkillToEnemy", "FaZhenSkillToSelf",
                                                "SkillToEnemyEffectToOther", "RandZoneSkillToSelfPos", "RectRangeSkillToEnemyPos" };

                    for (int i = 0; i < group_name_list.Length; i++)
                    {
                        XElement group_node = new XElement(group_name_list[i]);
                        root_node.Add(group_node);
                        group_dic.Add(group_name_list[i], group_node);
                    }
                }

                foreach (List<object> val_list in table.itemList)
                {
                    int skill_id = Convert.ToInt32(val_list[0]);
                    string group_name = this.GetSkillGroupName(skill_id);

                    XElement group_node = null;
                    if (!group_dic.TryGetValue(group_name, out group_node))
                    {
                        group_node = new XElement(group_name);
                        root_node.Add(group_node);
                        group_dic.Add(group_name, group_node);
                    }

                    XElement path_node = new XElement("skill");
                    path_node.SetValue(string.Format("monsterskills/{0}{1}.xml", group_name, skill_id));
                    group_node.Add(path_node);
                }

                string path = string.Format("{0}/gameworld/skill/MonsterPetSkillManager.xml", ConfigIni.XmlDir);

                Command.Instance.AddSvnAddFilePath(path);
                Writer.Instance.WriteXml(path, doc);

                Command.Instance.AddSvnCommitFilePath(path);
            }

            return true;
        }

        private string GetSkillGroupName(int skill_id)
        {
            string group_name = string.Empty;

            if (skill_id > 10000 && skill_id < 11000)
            {
                group_name = "CommonSkillToEnemy";
            }
            else if (skill_id > 11000 && skill_id < 12000)
            {
                group_name = "RangeCommonSkillToEnemyPos";
            }
            else if (skill_id > 12000 && skill_id < 13000)
            {
                group_name = "CommonSkillToSelf";
            }
            else if (skill_id > 13000 && skill_id < 14000)
            {
                group_name = "RangeCommonSkillToSelfPos";
            }
            else if (skill_id > 14000 && skill_id < 15000)
            {
                group_name = "FaZhenSkillToSelf";
            }
            else if (skill_id > 15000 && skill_id < 16000)
            {
                group_name = "FaZhenSkillToEnemy";
            }
            else if (skill_id > 16000 && skill_id < 17000)
            {
                group_name = "SkillToEnemyEffectToOther";
            }
            else if (skill_id > 17000 && skill_id < 18000)
            {
                group_name = "RandZoneSkillToSelfPos";
            }
            else if (skill_id > 18000 && skill_id < 19000)
            {
                group_name = "RectRangeSkillToEnemyPos";
            }
            else if (skill_id >= 20000 || skill_id < 5000)
            {
                group_name = "CommonSkillToEnemy";
            }

            return group_name;
        }
    }
}
