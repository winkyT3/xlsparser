using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using MySql.Data.MySqlClient;
using System.IO;
using System.Linq;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace xlsparser
{
    class Command : Singleton<Command>
    {
        public struct Order
        {
            public string cmd;
            public CMD_TYPE cmdType;
            public string log;
        }

        public enum CMD_TYPE
        {
            NONE = 0,
            SVN_UPDATE,
            SVN_COMMIT,
            SVN_ADD,
            SVN_CLEANUP,
            SVN_REVERT,
            IDENTIFY,
            HOTUPDTEA,
        }

        public delegate void Output(string line_data, Color color);

        private Process process;
        private Output dataReceiveds;
        private Queue<Order> orderQueue = new Queue<Order>();
        private CMD_TYPE executeCmdType;
        private List<string> conflictList = new List<string>();

        private HashSet<string> svnAddFilePaths = new HashSet<string>();
        private HashSet<string> svnCommitFilePaths = new HashSet<string>();

        public Command()
        {
            this.process = new System.Diagnostics.Process();
            this.process.StartInfo.FileName = "cmd.exe";
            this.process.StartInfo.UseShellExecute = false;
            this.process.StartInfo.RedirectStandardInput = true;
            this.process.StartInfo.RedirectStandardOutput = true;
            this.process.StartInfo.RedirectStandardError = true;
            this.process.EnableRaisingEvents = true;
            this.process.StartInfo.CreateNoWindow = true;

            this.process.OutputDataReceived += new DataReceivedEventHandler(this.OnDataReceived);
            this.process.ErrorDataReceived += new DataReceivedEventHandler(this.OnErrorDataReceived);
            this.process.Exited += new EventHandler(this.OnExit);

            this.process.Start();
            this.process.BeginOutputReadLine();
        }

        public void ListenDataReceivedEvent(Output callback)
        {
            this.dataReceiveds += callback;
        }

        public void AddSvnAddFilePath(string path)
        {
            if (this.svnAddFilePaths.Contains(path))
            {
                return;
            }

            if (File.Exists(path))
            {
                return;
            }

            this.svnAddFilePaths.Add(path);
        }

        public void AddSvnCommitFilePath(string path)
        {
            if (this.svnCommitFilePaths.Contains(path))
            {
                return;
            }

            this.svnCommitFilePaths.Add(path);
        }

        public void OpenExcel(string path)
        {
            this.PrintLog("打开excel文件:" + path, Color.Blue);
            Process.Start(path);
        }

        public void SvnUp()
        {
            if (CMD_TYPE.NONE != this.executeCmdType)
            {
                this.PrintLog("正在执行命令，别心急...", Color.Red);
                return;
            }

            this.PrintLog("开始SVN更新，请稍等...", Color.Blue);
            this.Execute("svn up ./ & exit", CMD_TYPE.SVN_UPDATE);

            this.Execute(string.Format("svn up {0} & exit", ConfigIni.XmlDir), CMD_TYPE.SVN_UPDATE);
            this.Execute(string.Format("svn up {0} & exit", ConfigIni.LuaDir), CMD_TYPE.SVN_UPDATE);
            this.Execute(string.Format("svn up {0} & exit", ConfigIni.XlsDir), CMD_TYPE.SVN_UPDATE);
        }

        public void SvnCommit(string log)
        {
            if (CMD_TYPE.NONE != this.executeCmdType)
            {
                this.PrintLog("正在执行命令，别心急...", Color.Red);
                return;
            }

            if (this.svnAddFilePaths.Count <= 0 && this.svnCommitFilePaths.Count <= 0)
            {
                this.PrintLog("没有可提交文件！请先点击生成配置", Color.Blue);
                return;
            }

            this.PrintLog("开始SVN提交，请稍等...", Color.Blue);

            // add lua, xml
            {
                foreach (string path in this.svnAddFilePaths)
                {
                    this.Execute(string.Format("svn add --force {0} & exit", path), CMD_TYPE.SVN_ADD);
                }
            }

            if (string.IsNullOrEmpty(log))
            {
                log = "commit by xlsparser";
            }

            Dictionary<char, List<string>> commit_file_group_paths = new Dictionary<char, List<string>>();
            foreach (string path in this.svnCommitFilePaths)
            {
                List<string> path_groups = null;
                if (!commit_file_group_paths.TryGetValue(path[0], out path_groups))
                {
                    path_groups = new List<string>();
                    commit_file_group_paths.Add(path[0], path_groups);
                }

                path_groups.Add(path);
            }

            foreach (var group_paths in commit_file_group_paths.Values)
            {
                string commit_files = string.Empty;
                int count = 0;
                foreach (string path in group_paths)
                {
                    ++count;
                    commit_files = commit_files + path + (count != group_paths.Count ? " " : "");
                }

                if (!string.IsNullOrEmpty(commit_files))
                {
                    this.Execute(string.Format("svn commit -m \"{0}\" {1} & exit", log, commit_files), CMD_TYPE.SVN_COMMIT, null);
                }
            }
        }

        public void SvnClearup()
        {
            if (CMD_TYPE.NONE != this.executeCmdType)
            {
                this.PrintLog("正在执行命令，别心急...", Color.Red);
                return;
            }

            this.PrintLog("开始SVN清理，请稍等...", Color.Blue);

            this.Execute(string.Format("svn cleanup {0} & exit", ConfigIni.XmlDir), CMD_TYPE.SVN_CLEANUP);
            this.Execute(string.Format("svn cleanup {0} & exit", ConfigIni.LuaDir), CMD_TYPE.SVN_CLEANUP);
            this.Execute(string.Format("svn cleanup {0} & exit", ConfigIni.XlsDir), CMD_TYPE.SVN_CLEANUP);
        }

        public void SvnRevert()
        {
            if (CMD_TYPE.NONE != this.executeCmdType)
            {
                this.PrintLog("正在执行命令，别心急...", Color.Red);
                return;
            }

            if (DialogResult.OK == MessageBox.Show("SVN还原操作将会删除你本地所有修改的文件，继续吗？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning))
            {
                this.PrintLog("开始SVN还原，请稍等...", Color.Blue);
                this.Execute(string.Format("svn revert -R {0} & exit", ConfigIni.XmlDir), CMD_TYPE.SVN_REVERT);
                this.Execute(string.Format("svn revert -R {0} & exit", ConfigIni.LuaDir), CMD_TYPE.SVN_REVERT);
                this.Execute(string.Format("svn revert -R {0} & exit", ConfigIni.XlsDir), CMD_TYPE.SVN_REVERT);
            }
        }

        public void Identify()
        {
            if (CMD_TYPE.NONE != this.executeCmdType)
            {
                this.PrintLog("正在执行命令，别心急...", Color.Red);
                return;
            }

            this.PrintLog("开始验证，请稍等...", Color.Blue);

            string rsakey_save_path = ConfigIni.GetRsaKeySavePath();
            Writer.Instance.WriteRsa(rsakey_save_path);

            try
            {
                string connect = string.Format("ssh 192.168.9.60 -p 18888 -i {0} -l root -o StrictHostKeyChecking=no",
                        ConfigIni.GetRsaKeySavePath());
                string cmd = string.Format("cd /data/game/{0}/htdocs/workbench/xls2xml", ConfigIni.ProjectName) +
                            string.Format("&&su - www -c 'svn up /data/game/{0}/htdocs/workbench/xls2xml/server/'", ConfigIni.ProjectName) +
                            string.Format("&&cp -f /{0}/workspace/publish_debug/EXEgameworld_debug ./check/", ConfigIni.ProjectName) +
                            "&&cd ./check" +
                            "&&./EXEgameworld_debug serverconfig.xml -checkres";

                string full_cmd = string.Format("{0} \"{1}\" & exit", connect, cmd);
                this.Execute(full_cmd, CMD_TYPE.IDENTIFY);
            }
            catch (Exception)
            {
                this.PrintLog("验证失败", Color.Red);
            }
        }

        public void HotUpdate(int index, string type_str)
        {
            if (CMD_TYPE.NONE != this.executeCmdType)
            {
                this.PrintLog("正在执行命令，别心急...", Color.Red);
                return;
            }

            {
                try
                {
                    string rsakey_save_path = ConfigIni.GetRsaKeySavePath();
                    Writer.Instance.WriteRsa(rsakey_save_path);

                    this.PrintLog("连接服务器开始热更，请稍等..." + type_str, Color.Blue);
                    string connect = string.Format("ssh 192.168.9.60 -p 18888 -i {0} -l root -o StrictHostKeyChecking=no",
                    rsakey_save_path);

                    string cmd = string.Format("svn up /{0}/workspace/config/gameworld/", ConfigIni.ProjectName);
                    string full_cmd = string.Format("{0} \"{1}\" & exit", connect, cmd);
                    this.Execute(full_cmd, CMD_TYPE.HOTUPDTEA);
                }
                catch (Exception)
                {
                    this.PrintLog("连接服务器失败", Color.Red);
                }
            }

            {
                string connect_str = ConfigIni.Database;
                MySqlConnection mysqlcon = new MySqlConnection(connect_str);
                try
                {
                    mysqlcon.Open();
                }
                catch (Exception)
                {
                    this.PrintLog("连接数据库失败" + connect_str, Color.Red);
                }

                try
                {
                    string cmd_str = string.Format("insert into command(creator,createtime,type,cmd,confirmtime)"
                                                    + "values(\"test\", {0}, 2, \"Cmd Reload {1} 0 0\", 0);", 1, index + 1);
                    MySqlCommand cmd = new MySqlCommand(cmd_str, mysqlcon);
                    int num = cmd.ExecuteNonQuery();
                    if (num > 0)
                    {
                        this.PrintLog("热更成功 " + type_str, Color.Green);
                    }
                    else
                    {
                        this.PrintLog("热更失败 " + type_str, Color.Red);
                    }
                }
                catch (Exception)
                {
                    this.PrintLog("热更失败 " + type_str, Color.Red);
                }

                mysqlcon.Close();
            }
        }

        private void Execute(string cmd, CMD_TYPE cmd_type, string log = null)
        {
            Order order = new Order();
            order.cmdType = cmd_type;
            order.cmd = cmd;
            order.log = log;

            if (CMD_TYPE.NONE != this.executeCmdType)
            {
                this.orderQueue.Enqueue(order);
            }
            else
            {
                this.DoExecute(order);
            }
        }

        private void DoExecute(Order order)
        {
            this.CloseProcess();

            if (!string.IsNullOrEmpty(order.log))
            {
                this.PrintLog(order.log, Color.Green);
            }

            this.executeCmdType = order.cmdType;
            this.process.Start();
            this.process.StandardInput.WriteLine(order.cmd);
            this.process.BeginOutputReadLine();
            this.process.BeginErrorReadLine();
            this.process.StandardInput.AutoFlush = true;
        }

        public void PrintLog(string log, Color color)
        {
            this.dataReceiveds(log, color);
        }

        private void OnDataReceived(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if (null == outLine.Data)
            {
                return;
            }

            bool is_log = true;
            if (outLine.Data.IndexOf(System.Windows.Forms.Application.StartupPath) >= 0
                || outLine.Data.IndexOf("Microsoft") >= 0)
            {
                is_log = false;
            }

            Color color = Color.Black;
            if (0 == outLine.Data.IndexOf("Adding") ||
                0 == outLine.Data.IndexOf("Sending") ||
                0 == outLine.Data.IndexOf("U ") ||
                0 == outLine.Data.IndexOf("Updating"))
            {
                color = Color.Green;
                is_log = true;
            }

            if (0 == outLine.Data.IndexOf("Reverted"))
            {
                color = Color.Purple;
                is_log = true;
            }

            if (0 == outLine.Data.IndexOf("C "))
            {
                this.conflictList.Add(outLine.Data);
                color = Color.Red;
                is_log = true;
            }

            if (is_log)
            {
                this.dataReceiveds(outLine.Data, color);
            }
        }

        private void OnErrorDataReceived(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if (null == outLine.Data)
            {
                return;
            }

            if (CMD_TYPE.IDENTIFY == this.executeCmdType ||
                CMD_TYPE.HOTUPDTEA == this.executeCmdType)
            {
                return;
            }

            this.CloseProcess();

            Writer.Instance.DelRsa(ConfigIni.GetRsaKeySavePath());

            Console.WriteLine(outLine.Data);
            this.dataReceiveds(outLine.Data, Color.Red);

            this.executeCmdType = CMD_TYPE.NONE;
            this.PrintLog("执行失败，中断操作", Color.Red);

            if (0 == outLine.Data.IndexOf("svn: E160028"))
            {
                if (DialogResult.OK == MessageBox.Show("本地文件过期，点确定进行SVN更新操作", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Error))
                {
                    this.SvnUp();
                }
            }
            else if (0 == outLine.Data.IndexOf("svn: E155037"))
            {
                if (DialogResult.OK == MessageBox.Show("SVN被锁定，点确定后进行SVN清理操作", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Error))
                {
                    this.SvnClearup();
                }
            }
            else if(0 == outLine.Data.IndexOf("svn: E155009"))
            {
                if (DialogResult.OK == MessageBox.Show("当前打开excel文件冲突，关闭文件后，再点击确定进行清理操作", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Error))
                {
                    this.SvnClearup();
                }
            }
            else if(0 == outLine.Data.IndexOf("svn: E"))
            {
                MessageBox.Show("SVN出现未知错误，请在本地手动解决", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnExit(object sender, EventArgs e)
        {
            if (CMD_TYPE.SVN_ADD == this.executeCmdType)
            {
                this.svnAddFilePaths.Clear();
            }

            if (CMD_TYPE.SVN_COMMIT == this.executeCmdType)
            {
                this.svnCommitFilePaths.Clear();
            }

            if (this.orderQueue.Count > 0)
            {
                this.DoExecute(this.orderQueue.Dequeue());
                return;
            }

            this.executeCmdType = CMD_TYPE.NONE;
            this.CloseProcess();
            Writer.Instance.DelRsa(ConfigIni.GetRsaKeySavePath());
            this.PrintLog("执行完成", Color.Blue);

            if (this.conflictList.Count > 0)
            {
                if (DialogResult.OK == MessageBox.Show("发现冲突文件，点确定后进行SVN还原。注意：还原将删除你本地修改的冲突文件", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Error))
                {
                    foreach (string path in this.conflictList)
                    {
                        this.Execute(string.Format("svn revert {0} & exit", path), CMD_TYPE.SVN_CLEANUP);
                    }
                    this.conflictList.Clear();
                }
            }
        }

        private void CloseProcess()
        {
            try
            {
                this.process.CancelOutputRead();
                this.process.CancelErrorRead();
                this.process.Close();
            }
            catch (Exception)
            {
            }
        }
    }
}
