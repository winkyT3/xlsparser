using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace xlsparser
{
    class LuaBuilder
    {
        public static void SimplifiedTableList(List<Table> table_list)
        {
            List<Table> default_table_list = new List<Table>();

            for (int i = 0; i < table_list.Count; ++i)
            {
                Table default_table = SimplifiedTable(table_list[i]);
                default_table_list.Add(default_table);
            }

            for (int i = 0; i < default_table_list.Count; ++i)
            {
                table_list.Add(default_table_list[i]);
            }

        }

        public static Table SimplifiedTable(Table table)
        {
            List<KeyT> default_key_list = new List<KeyT>();
            List<Dictionary<object, int>> value_count_list = new List<Dictionary<object, int>>();

            // init key
            for (int j = 0; j < table.keyList.Count; ++ j)
            {
                value_count_list.Add(new Dictionary<object, int>());

                KeyT key_T = table.keyList[j].Clone();
                if (KEY_TYPE.MAIN_KEY == key_T.keyType)
                {
                    key_T.keyType = KEY_TYPE.NORMA;
                }

                default_key_list.Add(key_T);
            }

            // calc value disappear count
            for(int m = 0; m < table.itemList.Count; ++ m)
            {
                List<object> value_list = table.itemList[m];
                for (int n = 0; n < value_list.Count; ++ n)
                {
                    Dictionary<object, int> count_dic = value_count_list[n];
                    object value = value_list[n];

                    if (count_dic.ContainsKey(value))
                    {
                        ++count_dic[value];
                    }
                    else
                    {
                        count_dic.Add(value, 1);
                    }
                }
            }

            // find default val
            List<object> default_value_list = new List<object>();
            for (int x = 0; x < value_count_list.Count; ++ x)
            {
                Dictionary<object, int> count_dic = value_count_list[x];
                int max_count = 0;
                object default_val = null;

                foreach (var item in count_dic)
                {
                    if (item.Value > max_count)
                    {
                        max_count = item.Value;
                        default_val = item.Key;
                    }
                }

                default_value_list.Add(default_val);
            }

            // flag del default val in itemlist
            for (int m = 0; m < default_value_list.Count; ++ m)
            {
                object default_val = default_value_list[m];
                for (int n = 0; n < table.itemList.Count; ++ n)
                {
                    if (default_val.Equals(table.itemList[n][m]) &&
                        KEY_TYPE.MAIN_KEY != table.keyList[m].keyType)
                    {
                        table.itemList[n][m] = null;
                    }
                }
            }

            // add default to memtable list
            Table default_table = new Table();
            default_table.name = string.Format("{0}_default_table", table.name);
            default_table.tableType = TABLE_TYPE.SIMPLE;
            default_table.keyList = default_key_list;
            default_table.itemList = new List<List<object>>();
            default_table.itemList.Add(default_value_list);

            return default_table;
        }

        public static string GetLuaTable(Table table)
        {
            if (!string.IsNullOrEmpty(table.outFlag) && 
                table.outFlag.IndexOf(ConfigIni.OUTPUT_S) >= 0)
            {
                return "";
            }

            int item_count = table.itemList.Count;
            int key_count = table.keyList.Count;
            if (key_count <= 0)
            {
                return "";
            }

            StringBuilder builder = new StringBuilder();
            if (TABLE_TYPE.NORMAL == table.tableType)
            {
                builder.Append(string.Format("{0}={{\n", table.name));
            }
            else if (TABLE_TYPE.SIMPLE == table.tableType)
            {
                builder.Append(string.Format("{0}=", table.name));
            }

            for (int i = 0; i < item_count; ++ i)
            {
                List<object> value_list = table.itemList[i];

                if (KEY_TYPE.MAIN_KEY ==  table.keyList[0].keyType)
                {
                    if (value_list[0].GetType() == typeof(string))
                    {
                        builder.Append(string.Format("[\"{0}\"]={{", value_list[0]));
                    }
                    else
                    {
                        builder.Append(string.Format("[{0}]={{", value_list[0]));
                    }
                   
                }
                else
                {
                    builder.Append("{");
                }
                
                for (int j = 0; j < key_count; ++j)
                {
                    KeyT key_T = table.keyList[j];

                    if (key_T.outFlag.IndexOf(ConfigIni.OUTPUT_C) < 0 ||
                        null == value_list[j])
                    {
                        continue;
                    }

                    builder.Append(GetToLuaKeyValue(key_T, value_list[j]));
                    builder.Append(",");
                }

                builder.Append(i != item_count - 1 ? "},\n" : "}" );
            }

            if (TABLE_TYPE.NORMAL == table.tableType)
            {
                builder.Append("}");
            }

            return builder.ToString();
        }

        private static string GetToLuaKeyValue(KeyT key_T, object value)
        {
            if (KEY_TYPE.ITEM == key_T.keyType)
            {
                return GetLuaItemTable(key_T.key, value.ToString());
            }
            else if (KEY_TYPE.ITEM_LIST == key_T.keyType)
            {
                return ConvertToLuaItemListTable(key_T.key, value.ToString());
            }
            else if  (KEY_TYPE.NPC_OBJ == key_T.keyType)
            {
                return GetNpcTable(key_T.key, value);
            }
            else if (KEY_TYPE.SCENE_OBJ_LIST == key_T.keyType)
            {
                return GetSceneObjListTable(key_T.key, value);
            }
            else if (KEY_TYPE.SPLIT_LIST == key_T.keyType)
            {
                return GetSplitListTable(key_T.key, value);
            }
            else if (KEY_TYPE.STEP_LIST == key_T.keyType)
            {
                return GetStepListTable(key_T.key, value);
            }
            else if (KEY_TYPE.NORMA == key_T.keyType || KEY_TYPE.MAIN_KEY == key_T.keyType)
            {
                if (value.GetType() == typeof(string))
                {
                    return string.Format("{0}=\"{1}\"", key_T.key, value);
                }
                else
                {
                    return string.Format("{0}={1}", key_T.key, value);
                }
            }
            else
            {
                return "error";
            }
        }

        private static string GetLuaItemTable(string key, string value)
        {
            string[] key_ary = key.Split(',');
            string[] val_ary = value.Split(':');
            if (val_ary.Length != 3)
            {
                return "error";
            }

            return string.Format("{0}={{item_id={1},num={2},is_bind={3}}}", key_ary[0], val_ary[0], val_ary[1], val_ary[2]);
        }

        private static string ConvertToLuaItemListTable(string key, string value)
        {
            string[] key_ary = key.Split(',');
            string[] list_ary = value.Split(',');
            if (list_ary.Length <= 0)
            {
                return "error";
            }
           
            string list_str = "";
            for (int i = 0; i < list_ary.Length; ++ i)
            {
                string[] val_ary = list_ary[i].Split(':');
                string tail = i != list_ary.Length - 1 ? "," : "";

                if (val_ary.Length != 3)
                {
                    list_str += "";
                }
                else
                {
                    list_str += string.Format("[{0}]={{item_id={1},num={2},is_bind={3}}}{4}", i, val_ary[0], val_ary[1], val_ary[2], tail);
                }
            }

            return string.Format("{0}={{{1}}}", key_ary[0], list_str);
        }

        private static string GetNpcTable(string key, object value)
        {
           List<SceneObjVo> list = SceneObjects.Instance.GetNpcList(StringUtil.ConvertToInt32(value.ToString()));
           if (null == list || list.Count <= 0)
            {
                return string.Format("{0}={{}}", key);
            }

            return string.Format("{0}={{id={1},scene={2},x={3},y={4}}}", key, list[0].id, list[0].sceneId, list[0].x, list[0].y);
        }

        private static string GetSceneObjListTable(string key, object value)
        {
            string[] ary = value.ToString().Split(',');
            if (ary.Length < 2)
            {
                return string.Format("{0}={{}}", key);
            }

            int obj_type = StringUtil.ConvertToInt32(ary[0]);
            int obj_id = StringUtil.ConvertToInt32(ary[1]);

            List<SceneObjVo> list = null;
            if (0 == obj_type)
            {
                list = SceneObjects.Instance.GetNpcList(obj_id);
            }

            if (1 == obj_type)
            {
                list = SceneObjects.Instance.GetMonsterList(obj_id);
            }

            if (2 == obj_type)
            {
                list = SceneObjects.Instance.GetGatherList(obj_id);
            }

            if (null == list || list.Count <= 0)
            {
                return string.Format("{0}={{}}", key);
            }

            string list_str = "";
            foreach (SceneObjVo vo in list)
            {
                list_str += string.Format("{{id={0},scene={1},x={2},y={3}}},", vo.id, vo.sceneId, vo.x, vo.y);
            }

            return string.Format("{0}={{{1}}}", key, list_str);
        }

        private static string GetSplitListTable(string key, object value)
        {
            string[] ary = value.ToString().Split(',');
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < ary.Length; i++)
            {
                builder.Append(string.Format("\"{0}\",", ary[i]));
            }

            return string.Format("{0}={{{1}}}", key, builder.ToString());
        }

        private static string GetStepListTable(string key, object value)
        {
            string[] kyes_ary = { "step_type", "step_param", "module_name",
                "ui_name", "ui_param", "arrow_dir", "arrow_tip",
                "is_rect_effect", "is_finger_effect", "is_modal", "unuseful", "offset_x", "offset_y"};
            string[] step_ary = value.ToString().Split(',');
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < step_ary.Length; i++)
            {
                StringBuilder step_builder = new StringBuilder();
                string[] values_ary = step_ary[i].Split(':');

                for (int j = 0; j < values_ary.Length; j++)
                {
                    if (Regex.Match(values_ary[j], "^[-+]?[0-9]*\\.?[0-9]+$").Success)
                    {
                        step_builder.Append(string.Format("{0}={1},", kyes_ary[j], values_ary[j]));
                    }
                    else
                    {
                        step_builder.Append(string.Format("{0}=\"{1}\",", kyes_ary[j], values_ary[j]));
                    }
                }

                builder.Append(string.Format("[{0}]={{{1}}},", i + 1, step_builder.ToString()));
            }

            return string.Format("{0}={{{1}}}", key, builder.ToString());
        }
    }
}
