using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;

namespace xlsparser
{
    struct SceneObjVo
    {
        public int id;
        public int sceneId;
        public int x;
        public int y;
    }

    struct SceneVo
    {
        public int sceneId;
        public int sceneType;
    }

    enum SceneObjType
    {
        NPC,
        MONSTER,
        GATHER,
    }

    class SceneObjects : Singleton<SceneObjects>
    {
        private Dictionary<int, List<SceneObjVo>> npcDic = new Dictionary<int, List<SceneObjVo>>();
        private Dictionary<int, List<SceneObjVo>> monsterDic = new Dictionary<int, List<SceneObjVo>>();
        private Dictionary<int, List<SceneObjVo>> gatherDic = new Dictionary<int, List<SceneObjVo>>();

        private List<SceneVo> sceneVoList = new List<SceneVo>();

        public List<SceneObjVo> GetNpcList(int npc_id)
        {
            List<SceneObjVo> list = null;
            this.npcDic.TryGetValue(npc_id, out list);

            return list;
        }

        public List<SceneObjVo> GetMonsterList(int monster_id)
        {
            List<SceneObjVo> list = null;
            this.monsterDic.TryGetValue(monster_id, out list);

            return list;
        }

        public List<SceneObjVo> GetGatherList(int gather_id)
        {
            List<SceneObjVo> list = null;
            this.gatherDic.TryGetValue(gather_id, out list);

            return list;
        }

        public bool ReadAllObjects()
        {
            this.npcDic.Clear();
            this.monsterDic.Clear();
            this.gatherDic.Clear();

            if (!this.ReadSceneIdList())
            {
                return false;
            }

            if (!this.ReadAllScene())
            {
                return false;
            }

            return true;
        }

        private bool ReadSceneIdList()
        {
            this.sceneVoList.Clear();

            string path = string.Format("{0}/../config_map.lua", ConfigIni.LuaDir);
            if (!File.Exists(path))
            {
                return false;
            }

            try 
	        {	        
		        string content = File.ReadAllText(path);
                MatchCollection match_list = Regex.Matches(content, "\\[(\\d*)\\].*sceneType = (\\d*)");
                foreach (Match m in match_list)
                {
                    if (m.Groups.Count >= 3)
                    {
                        SceneVo vo = new SceneVo();
                        vo.sceneId = Convert.ToInt32(m.Groups[1].ToString());
                        vo.sceneType = Convert.ToInt32(m.Groups[2].ToString());
          
                        this.sceneVoList.Add(vo);
                    }
                }

                return true;
	        }
	        catch (Exception)
	        {
		        return false;
		        throw;
	        }
        }

        private bool ReadAllScene()
        {
            foreach (SceneVo scene_vo in this.sceneVoList)
            {
                string path = string.Format("{0}/../scenes/scene_{1}.lua", ConfigIni.LuaDir, scene_vo.sceneId);
                if (!File.Exists(path))
                {
                    continue;
                }

                try
                {
                    string content = File.ReadAllText(path);
                    this.ReadObjsInConfig(scene_vo, content, "npcs = \\{([\\s\\S]*)\\}[\\s\\S]*monsters = ", this.npcDic, SceneObjType.NPC);
                    this.ReadObjsInConfig(scene_vo, content, "monsters = \\{([\\s\\S]*)\\}[\\s\\S]*doors = ", this.monsterDic, SceneObjType.MONSTER);
                    this.ReadObjsInConfig(scene_vo, content, "gathers = \\{([\\s\\S]*)\\}[\\s\\S]*jumppoints = ", this.gatherDic, SceneObjType.GATHER);
                }
                catch (Exception)
                {
                    return false;
                    throw;
                }
            }

            return true;
        }

        private void ReadObjsInConfig(SceneVo scene_vo, string scene_content, string pattern, Dictionary<int, List<SceneObjVo>> dic, SceneObjType obj_type)
        {
            Match match = Regex.Match(scene_content, pattern);
            if (match.Groups.Count < 2)
            {
                return;
            }

            bool is_first = true;
            Match m = Regex.Match(match.Groups[1].ToString(), "id=(\\d*), x=(\\d*), y=(\\d*)");
            while (m.Success)
            {
                if (m.Groups.Count >= 4)
                {
                    SceneObjVo vo = new SceneObjVo();
                    vo.sceneId = scene_vo.sceneId;
                    vo.id = Convert.ToInt32(m.Groups[1].ToString());
                    vo.x = Convert.ToInt32(m.Groups[2].ToString());
                    vo.y = Convert.ToInt32(m.Groups[3].ToString());

                    // 普通场景的限制
                    if (0 == scene_vo.sceneType)
                    {
                        if (SceneObjType.NPC == obj_type && dic.ContainsKey(vo.id))
                        {
                            Command.Instance.PrintLog(string.Format("错误：有相同的NPC,可能导致任务寻路失败, scene_id={0}, npc_id = {1}", vo.sceneId, vo.id), Color.Red);
                        }

                        if (is_first)
                        {
                            is_first = false;
                            if (SceneObjType.MONSTER == obj_type && dic.ContainsKey(vo.id))
                            {
                                Command.Instance.PrintLog(string.Format("警告：不同场景有相同的怪物,可能导致任务寻路失败, scene_id={0}, monster_id = {1}", vo.sceneId, vo.id), Color.YellowGreen);
                            }

                            if (SceneObjType.GATHER == obj_type && dic.ContainsKey(vo.id))
                            {
                                Command.Instance.PrintLog(string.Format("错误：不同场景有相同的的采集物,可能导致任务寻路失败, scene_id={0}, gather_id = {1}", vo.sceneId, vo.id), Color.Red);
                            }
                        }
                    }
                    
                    List<SceneObjVo> vo_list;
                    if (!dic.TryGetValue(vo.id, out vo_list))
                    {
                        vo_list = new List<SceneObjVo>();
                        dic.Add(vo.id, vo_list);
                    }

                    vo_list.Add(vo);
                }

                m = m.NextMatch();
            }
        }
    }
}
