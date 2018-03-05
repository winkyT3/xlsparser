using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using xlsparser;

namespace xlsparser.view
{
    public struct XlsItem
    {
        public string showName;
        public string path;
        public string outputFlag;
        public XLS_PARSER_TYPE parser_type;

        public XlsItem(string _showName, string _path, string _outputFlag, XLS_PARSER_TYPE _parser_type = XLS_PARSER_TYPE.NORMAL)
        {
            showName = _showName;
            path = _path;
            outputFlag = _outputFlag;
            parser_type = _parser_type;
        }
    }

    public partial class BuildWin : Form
    {
        private TextBox serarchText;
        private TextBox logText;
        private ComboBox hotupdateCombobox;
        private RichTextBox consoleText;

        private FlowLayoutPanel contentLayout;
        private FlowLayoutPanel topLayout;
        private FlowLayoutPanel searchLayout;
        private FlowLayoutPanel itemLayout;

        private List<XlsItem> xlsList = new List<XlsItem>();
        private Queue<XlsItemLayout> xlsItemLayoutQueue = new Queue<XlsItemLayout>();

        public BuildWin()
        {
            InitializeComponent();
            this.Text = string.Format("配置生成功具({0})", ConfigIni.ProjectName);
            this.Size = new Size(ConfigIni.WinWidth, ConfigIni.WinHeight);

            this.ReadItemContentXls();
            this.ReadMonsterContentXls();
            if (!this.ReadNormalContentsXls())
            {
                MessageBox.Show("读取通用配置目录失败");
                return;
            }

            this.CreateContents();
            Command.Instance.ListenDataReceivedEvent(this.OnRecevieData);
        }

        private void ReadItemContentXls()
        {
            xlsList.Add(new XlsItem("物品索引", string.Format("{0}/物品表/", ConfigIni.XlsDir), "s", XLS_PARSER_TYPE.ITEM_INDEX));
            xlsList.Add(new XlsItem("装备类", string.Format("{0}/物品表/W-装备.xls", ConfigIni.XlsDir), "cs", XLS_PARSER_TYPE.ITEM));
            xlsList.Add(new XlsItem("被动消耗类", string.Format("{0}/物品表/W-被动消耗类.xls", ConfigIni.XlsDir), "cs", XLS_PARSER_TYPE.ITEM));
            xlsList.Add(new XlsItem("主动使用消耗类", string.Format("{0}/物品表/W-主动使用消耗类.xls", ConfigIni.XlsDir), "cs", XLS_PARSER_TYPE.ITEM));
            xlsList.Add(new XlsItem("礼包类", string.Format("{0}/物品表/W-礼包类.xls", ConfigIni.XlsDir), "cs", XLS_PARSER_TYPE.ITEM));
            xlsList.Add(new XlsItem("虚拟类", string.Format("{0}/物品表/W-虚拟类.xls", ConfigIni.XlsDir), "c", XLS_PARSER_TYPE.ITEM));
        }

        private void ReadMonsterContentXls()
        {
            xlsList.Add(new XlsItem("G-怪物", string.Format("{0}/G-怪物.xls", ConfigIni.XlsDir), "cs", XLS_PARSER_TYPE.MONSTER));
            xlsList.Add(new XlsItem("G-怪物技能", string.Format("{0}/G-怪物技能.xls", ConfigIni.XlsDir), "cs", XLS_PARSER_TYPE.MONSTER_SKILL));
        }

        private bool ReadNormalContentsXls()
        {
            List<ISheet> sheet_list = new List<ISheet>();
            string path = string.Format("{0}/{1}.xls", ConfigIni.XlsDir, "填表说明/通用配置");
            if (!XlsReader.Instance.ReadExcel(path, sheet_list))
            {
                return false;
            }

            if (sheet_list.Count <= 0)
            {
                return false;
            }

            ISheet sheet = sheet_list[0];
            int row_num = 10000;
            for (int i = 2; i < row_num; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (null == row || row.LastCellNum < 3)
                {
                    break;
                }

                XlsItem xls_item = new XlsItem();
                xls_item.showName = row.GetCell(0).ToString();
                xls_item.path = string.Format("{0}/{1}.xls", ConfigIni.XlsDir, row.GetCell(1).ToString());
                xls_item.outputFlag = row.GetCell(2).ToString();
                this.FixXlsParserType(ref xls_item);

                this.xlsList.Add(xls_item);
            }

            return true;
        }

        private void FixXlsParserType(ref XlsItem xls_item)
        {
            if (xls_item.showName.Equals("D-掉落"))
            {
                xls_item.parser_type = XLS_PARSER_TYPE.DROP;
            }

            else if (xls_item.showName.Equals("B-BOSS"))
            {
                xls_item.parser_type = XLS_PARSER_TYPE.BOSS_SKILL_CONDITION;
            }

            else if (xls_item.showName.Equals("R-任务"))
            {
                xls_item.parser_type = XLS_PARSER_TYPE.TASK;
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            this.RefreshLayout(this.Size.Width, this.Size.Height);
        }

        private void RefreshLayout(int width, int height)
        {
            if (null != this.contentLayout)
            {
                this.contentLayout.Size = new Size(width - 25, height);
            }

            if (null != this.topLayout)
            {
                this.topLayout.Size = new Size(width, 40);
            }

            if (null != this.consoleText)
            {
                this.consoleText.Size = new Size(width - 25, 350);
            }

            if (null != this.searchLayout)
            {
                this.searchLayout.Size = new Size(width, 25);
            }

            if (null != this.itemLayout)
            {
                this.itemLayout.Size = new Size(width - 25, height - 475);
            }
        }

        private void CreateContents()
        {
            this.contentLayout = new FlowLayoutPanel();
            this.Controls.Add(this.contentLayout);

            // top
            {
                this.topLayout = new FlowLayoutPanel();
                this.contentLayout.Controls.Add(this.topLayout);

                Button svn_update_btn = new Button();
                svn_update_btn.Text = "SVN更新";
                this.topLayout.Controls.Add(svn_update_btn);
                svn_update_btn.Click += (object sender, EventArgs e) => { Command.Instance.SvnUp(); };

                Button svn_commit_btn = new Button();
                svn_commit_btn.Text = "SVN提交";
                this.topLayout.Controls.Add(svn_commit_btn);
                svn_commit_btn.Click += (object sender, EventArgs e) => {
                    Command.Instance.SvnCommit(this.logText.Text);
                    this.logText.Text = string.Empty;
                };

                this.logText = new TextBox();
                this.logText.Size = new Size(300, 40);
                this.topLayout.Controls.Add(this.logText);

                Button svn_clear_btn = new Button();
                svn_clear_btn.Text = "SVN清理";
                this.topLayout.Controls.Add(svn_clear_btn);
                svn_clear_btn.Click += (object sender, EventArgs e) => { Command.Instance.SvnClearup(); };

                Button svn_revert_btn = new Button();
                svn_revert_btn.Text = "SVN还原";
                this.topLayout.Controls.Add(svn_revert_btn);
                svn_revert_btn.Click += (object sender, EventArgs e) => { Command.Instance.SvnRevert(); };

                Button identify_btn = new Button();
                identify_btn.Text = "服务器验证";
                this.topLayout.Controls.Add(identify_btn);
                identify_btn.Click += (object sender, EventArgs e) => { Command.Instance.Identify(); };

                Button hotupdate_btn = new Button();
                hotupdate_btn.Text = "服务器热更";
                this.topLayout.Controls.Add(hotupdate_btn);
                hotupdate_btn.Click += (object sender, EventArgs e) => 
                {
                    Command.Instance.HotUpdate(this.hotupdateCombobox.SelectedIndex, this.hotupdateCombobox.SelectedValue.ToString());
                };

                {
                    List<string> source = new List<string>();
                    source.Add("全局");
                    source.Add("技能");
                    source.Add("任务");
                    source.Add("怪物");
                    source.Add("物品");
                    source.Add("逻辑");
                    source.Add("掉落");
                    source.Add("时装");
                    source.Add("场景");
                    this.hotupdateCombobox = new ComboBox();
                    this.hotupdateCombobox.DropDownStyle = ComboBoxStyle.DropDownList;
                    this.hotupdateCombobox.DataSource = source;
                    this.topLayout.Controls.Add(this.hotupdateCombobox);
                }
            }

            // console
            {
                this.consoleText = new RichTextBox();
                this.consoleText.Multiline = true;
                this.consoleText.WordWrap = false;
                this.consoleText.ScrollBars = RichTextBoxScrollBars.Both;
                this.consoleText.Text = "版本所有(c)2018 zhiwen。保留所有权限。";
                this.contentLayout.Controls.Add(this.consoleText);
            }

            // search
            {
                this.searchLayout = new FlowLayoutPanel();
                this.contentLayout.Controls.Add(this.searchLayout);

                Label label = new Label();
                label.Text = "搜索:";
                label.TextAlign = ContentAlignment.MiddleLeft;
                label.Width = 40;
                this.searchLayout.Controls.Add(label);

                this.serarchText = new TextBox();
                this.serarchText.Size = new Size(250, 40);
                this.serarchText.TextChanged += this.OnSearchTextChanged;
                this.searchLayout.Controls.Add(this.serarchText);
            }

            // bottom
            {
                this.itemLayout = new FlowLayoutPanel();
                this.itemLayout.AutoScroll = true;
                this.itemLayout.MouseEnter += (object sender, EventArgs e) => {
                    this.itemLayout.Focus();
                };
                this.contentLayout.Controls.Add(this.itemLayout);
                this.itemLayout.Focus();
            }

            this.RefreshLayout(ConfigIni.WinWidth, ConfigIni.WinHeight);

            this.RefreshXlsItemShow(this.xlsList);
        }

        private void OnRecevieData(string line_data, Color color)
        {
            this.BeginInvoke(new Action(() =>
            {
                if (null != line_data)
                {
                    this.PrintLog(line_data, color);
                }
            }));
        }

        private void PrintLog(string log, Color color)
        {
            this.consoleText.SelectionColor = color;
            this.consoleText.AppendText(log + "\n");
            this.consoleText.ScrollToCaret();
        }

        private void OnSearchTextChanged(object sender, EventArgs e)
        {
            string serach = this.serarchText.Text.Trim();
            if (string.IsNullOrEmpty(serach))
            {
                this.RefreshXlsItemShow(this.xlsList);
                return;
            }

            List<XlsItem> search_list = new List<XlsItem>();
            for (int i = 0; i < this.xlsList.Count; ++ i)
            {
                if (this.xlsList[i].showName.IndexOf(serach.ToUpper()) >= 0 ||
                    this.xlsList[i].showName.IndexOf(serach.ToLower()) >= 0)
                {
                    search_list.Add(this.xlsList[i]);
                }
            }

            if (search_list.Count > 0)
            {
                this.RefreshXlsItemShow(search_list);
            }
        }

        private void RefreshXlsItemShow(List<XlsItem> list)
        {
            this.itemLayout.Controls.Clear();

            Queue<XlsItemLayout> add_queue = new Queue<XlsItemLayout>();
            for (int i = 0; i < list.Count; ++i)
            {
                XlsItemLayout item = null;
                if (this.xlsItemLayoutQueue.Count > 0)
                {
                    item = this.xlsItemLayoutQueue.Dequeue();
                }
                else
                {
                    item = new XlsItemLayout();
                }
               
                item.SetData(list[i]);
                this.itemLayout.Controls.Add(item);
                add_queue.Enqueue(item);
            }

            foreach (var item in add_queue)
            {
                this.xlsItemLayoutQueue.Enqueue(item);
            }
        }
    }

    class XlsItemLayout : FlowLayoutPanel
    {
        private XlsItem xls_item;
        private LinkLabel xls_name_txt;
        private Button c_btn;
        private Button s_btn;

        public XlsItemLayout()
        {
            this.Size = new Size(300, 30);

            this.xls_name_txt = new LinkLabel();
            this.xls_name_txt.Width = 120;
            this.xls_name_txt.LinkColor = Color.Black;
            this.xls_name_txt.Height = this.Height;
            this.xls_name_txt.LinkBehavior = LinkBehavior.NeverUnderline;
            this.xls_name_txt.TextAlign = ContentAlignment.MiddleLeft;
            this.xls_name_txt.LinkClicked += this.OnClickLink;

            this.Controls.Add(this.xls_name_txt);

            this.c_btn = new Button();
            this.c_btn.Text = "Client";
            this.Controls.Add(this.c_btn);
            this.c_btn.Click += this.OnBuildClientCfg;

            this.s_btn = new Button();
            this.s_btn.Text = "Server";
            this.Controls.Add(this.s_btn);
            this.s_btn.Click += this.OnBuildServerCfg;
        }

        public void SetData(XlsItem xls_item)
        {
            this.xls_item = xls_item;
            
            xls_name_txt.Text = xls_item.showName;
            this.c_btn.Enabled = xls_item.outputFlag.IndexOf('c') >= 0;
            this.s_btn.Enabled = xls_item.outputFlag.IndexOf('s') >= 0;
        }

        private void OnClickLink(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Command.Instance.OpenExcel(xls_item.path);
        }

        private void OnBuildClientCfg(object sender, EventArgs e)
        {
             List<ISheet> sheet_list = new List<ISheet>();
             if (!XlsReader.Instance.ReadExcel(this.xls_item.path, sheet_list))
             {
                Command.Instance.PrintLog("读取Excel失败", Color.Red);
                return;
             }

            if (!Builder.Instance.BuildClient(this.xls_item.parser_type, sheet_list))
            {
                Command.Instance.PrintLog("生成配置失败", Color.Red);
                return;
            }

            this.OnBuildCfgSucc();
            Command.Instance.PrintLog(string.Format("生成配置成功： {0}", this.xls_item.showName), Color.Blue);
        }

        private void OnBuildServerCfg(object sender, EventArgs e)
        {
            if (XLS_PARSER_TYPE.ITEM_INDEX == this.xls_item.parser_type)
            {
                if (!Builder.Instance.BuildItemIndex(this.xls_item.path))
                {
                    Command.Instance.PrintLog("生成配置失败", Color.Red);
                    return;
                }

                Command.Instance.PrintLog("生成配置成功", Color.Blue);
            }
            else
            {
                List<ISheet> sheet_list = new List<ISheet>();
                if (!XlsReader.Instance.ReadExcel(this.xls_item.path, sheet_list))
                {
                    Command.Instance.PrintLog("读取Excel失败", Color.Red);
                    return;
                }

                if (!Builder.Instance.BuildServer(this.xls_item.parser_type, sheet_list))
                {
                    Command.Instance.PrintLog("生成配置失败", Color.Red);
                    return;
                }

                this.OnBuildCfgSucc();
                Command.Instance.PrintLog(string.Format("生成配置成功： {0}", this.xls_item.showName), Color.Blue);
            }

        }

        private void OnBuildCfgSucc()
        {
            Command.Instance.AddSvnAddFilePath(this.xls_item.path);
            Command.Instance.AddSvnCommitFilePath(this.xls_item.path);
        }
    }
}
