using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace ExcelExportTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitContent();
        }
        void InitContent() {
            this.MaximizeBox = false;

            Screen scr = Screen.PrimaryScreen;

            int width = scr.Bounds.Width / 2>900?scr.Bounds.Width/2:900;
            int height = scr.Bounds.Height / 2 / 2 > 500 ? scr.Bounds.Height / 2 : 500;

            this.Size = new Size(width,height);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Excel翻译表导出";
            //0
            Label label1 = new Label();
            label1.Location=new Point(24,8);
            label1.Size = new Size(500, 20);
            
            label1.Text = "请选择要用于生成翻译表的原excel表所在目录";

            this.Controls.Add(label1);

            //1
            TextBox txtbox = new TextBox();
            txtbox.Location = new Point(24, 25);
            txtbox.Size = new Size(500, 20);
            txtbox.Text = @"E:\zzhx\trunk\data\ai";
            this.Controls.Add(txtbox);

            //2
            Button btn1 = new Button();
            btn1.Location = new Point(530, 25);
            btn1.Size = new Size(100, 20);
            btn1.Text = "选择路径";
            this.Controls.Add(btn1);
            btn1.Click += btn1_Click;



            //ListBox listbox = new ListBox();
            //listbox.Location = new Point(24, 50);
            //listbox.Size = new Size(300, 300);
            //this.Controls.Add(listbox);
            //string[] translatetype = new string[] { "ch", "en", "ru", "ls", "ch", "en", "ru", "ls" };
            //List<string> translatetype = new List<string>();
            //parseXml(ref translatetype);

            //3
            GroupBox gbox = new GroupBox();
            gbox.Location = new Point(24, 100);
            gbox.Size = new Size(300, 200);

            Label label_bg = new Label();
            label_bg.Location = new Point(24, 30);
            label_bg.Size = new Size(200, 20);
            label_bg.Text = "选择要添加的语言种类";
            gbox.Controls.Add(label_bg);

            //int preIndex = 0;
            //int curIndex = 0;
            //for (int i = 0; i < translatetype.Count; i++)
            //{
            //    CheckBox cbox = new CheckBox();
            //    curIndex++;
            //    if (curIndex % 2 == 0)
            //        preIndex++;
            //    cbox.Location = new Point(24 + i % 2 * 50, 50 + preIndex * 30);
            //    cbox.Size = new Size(200, 30);
            //    cbox.Text = translatetype[i];
            //    gbox.Controls.Add(cbox);
            //}

            this.Controls.Add(gbox);
            //4
            Button btn_del = new Button();
            btn_del.Location = new Point(690, 330);
            btn_del.Size = new Size(90, 20);
            btn_del.Text = "删除翻译表";
            this.Controls.Add(btn_del);
            btn_del.Click += btn_translate_Del;
            
            //5
            Button btn_translate = new Button();
            btn_translate.Location = new Point(530, 150);
            btn_translate.Size = new Size(100, 20);
            btn_translate.Text = "导出Excel";
            this.Controls.Add(btn_translate);
            btn_translate.Click += btn_translate_Click;

            //6
            Label label_del = new Label();
            label_del.Location = new Point(24, 310);
            label_del.Size = new Size(500, 20);
            label_del.Text = "请选择要删除翻译表起始目录";
            this.Controls.Add(label_del);

            //7
            TextBox txtbox_del = new TextBox();
            txtbox_del.Location = new Point(24, 330);
            txtbox_del.Size = new Size(500, 20);
            txtbox_del.Text = @"E:\zzhx\trunk\data";
            this.Controls.Add(txtbox_del);

            //8
            Button btn_del_trans = new Button();
            btn_del_trans.Location = new Point(560, 330);
            btn_del_trans.Size = new Size(130, 20);
            btn_del_trans.Text = "选择删除翻译表路径";
            this.Controls.Add(btn_del_trans);
            btn_del_trans.Click += btn_del_trans_Click;

            //9
            Label label_filter = new Label();
            label_filter.Location = new Point(24, 50);
            label_filter.Size = new Size(500, 20);
            label_filter.Text = "请选择要过滤的目录";
            this.Controls.Add(label_filter);

            //10
            TextBox txtbox_filter = new TextBox();
            txtbox_filter.Location = new Point(24, 67);
            txtbox_filter.Size = new Size(500, 20);
            txtbox_filter.Text = @"";
            this.Controls.Add(txtbox_filter);

            //11
            Button btn_filter = new Button();
            btn_filter.Location = new Point(530, 67);
            btn_filter.Size = new Size(130, 20);
            btn_filter.Text = "选择过滤目录路径";
            this.Controls.Add(btn_filter);
            btn_filter.Click += btn_filter_Click;

            //12
            Label label_lan = new Label();
            label_lan.Location = new Point(24, 360);
            label_lan.Size = new Size(500, 20);
            label_lan.Text = "请选择语言配置文件";
            this.Controls.Add(label_lan);

            //13
            TextBox txtbox_lan = new TextBox();
            txtbox_lan.Location = new Point(24, 380);
            txtbox_lan.Size = new Size(500, 20);
            txtbox_lan.Text = @"";
            this.Controls.Add(txtbox_lan);

            //14
            Button btn_lan = new Button();
            btn_lan.Location = new Point(530, 380);
            btn_lan.Size = new Size(130, 20);
            btn_lan.Text = "选择语言配置路径";
            this.Controls.Add(btn_lan);
            btn_lan.Click += btn_lan_Click;

            //15
            Label label_lua = new Label();
            label_lua.Location = new Point(24, 410);
            label_lua.Size = new Size(500, 20);
            label_lua.Text = "请选择lua翻译字段名数据文件路径";
            this.Controls.Add(label_lua);

            //16
            TextBox txtbox_lua = new TextBox();
            txtbox_lua.Location = new Point(24, 430);
            txtbox_lua.Size = new Size(500, 20);
            txtbox_lua.Text = System.Environment.CurrentDirectory;//@"E:\zzhx\trunk\client\Tool\ExcelExportTool-master";
            this.Controls.Add(txtbox_lua);

            //17
            Button btn_lua = new Button();
            btn_lua.Location = new Point(530, 430);
            btn_lua.Size = new Size(130, 20);
            btn_lua.Text = "选择语言配置路径";
            this.Controls.Add(btn_lua);
            btn_lua.Click += btn_lua_Click;

            //18
            Button btn_refresh_lan = new Button();
            btn_refresh_lan.Location = new Point(660, 380);
            btn_refresh_lan.Size = new Size(130, 20);
            btn_refresh_lan.Text = "刷新语言配置";
            this.Controls.Add(btn_refresh_lan);
            btn_refresh_lan.Click += btn_refresh_lan_Click;
        }


        //-------------------------------------------------------------

        void UpdateTranslationType()
        {
            GroupBox gbox = this.Controls[3] as GroupBox;
            gbox.Controls.Clear();
            TextBox txtbox = this.Controls[13] as TextBox;
            List<string> translatetype = new List<string>();
            if (!parseXml(ref translatetype, txtbox.Text)) {
                return;
            }

            int preIndex = 0;
            int curIndex = 0;
            for (int i = 0; i < translatetype.Count; i++)
            {
                CheckBox cbox = new CheckBox();
                curIndex++;
                if (curIndex % 2 == 0)
                    preIndex++;
                cbox.Location = new Point(24 + i % 2 * 50, 50 + preIndex * 30);
                cbox.Size = new Size(200, 30);
                cbox.Text = translatetype[i];
                cbox.Checked = true;
                gbox.Controls.Add(cbox);
            }
        }

        void btn_lan_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            DialogResult result = openfile.ShowDialog();
            if (result == DialogResult.OK)
            {
                TextBox txtbox = this.Controls[13] as TextBox;
                txtbox.Text = openfile.FileName;
                UpdateTranslationType();
            }
        }

        /// <summary>
        /// 刷新语言翻译类型按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
         void btn_refresh_lan_Click(object sender, EventArgs e)
        {
            UpdateTranslationType();
        }

        bool parseXml(ref List<string> list, string xmlpath)
        {
            if (!File.Exists(xmlpath))
            {
                ShowTips("不存在" + xmlpath + "此xml文件");
                return false;
            }
            FileInfo fi = new FileInfo(xmlpath);
            if (fi.Extension != ".xml" || fi.Name != "language.xml") {
                ShowTips(xmlpath + "不是翻译类型的xml配置文件");
                return false;
            }
            list.Clear();
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(xmlpath);//@"..\..\..\language.xml"
            XmlElement root = xmldoc.DocumentElement;
            XmlNodeList listnodes = root.SelectNodes("/languagetype/language/type");
            foreach (XmlNode node in listnodes)
            {
                list.Add(node.InnerText);
            }
            return true;
        }

        /// <summary>
        /// 选择导出excel目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btn1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            DialogResult result= openfile.ShowDialog();
            if (result == DialogResult.OK) {
                TextBox txtbox= this.Controls[1] as TextBox;
                txtbox.Text = openfile.FileName;
                int index= txtbox.Text.LastIndexOf(@"\");
                txtbox.Text = txtbox.Text.Substring(0, index);
                //ShowTips(txtbox.Text);
            }
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btn_translate_Click(object sender, EventArgs e) {
            TextBox txtbox = this.Controls[1] as TextBox;
            TextBox luatxtbox=this.Controls[16] as TextBox;
            GroupBox gbox = this.Controls[3] as GroupBox;
            List<string> translateType = new List<string>();
            foreach (var item in gbox.Controls)
            {
                CheckBox cbox = item as CheckBox;
                if (cbox == null)
                    continue;
                if (cbox.Checked)
                {
                    translateType.Add(cbox.Text);
                }
            }
            TextBox txtbox_filter=this.Controls[10] as TextBox;
            //string[] translateTypeArr = translateType.ToArray();
            if (translateType.Count <= 0)
            {
                ShowTips("请添加语言再导出翻译表");
                return;
            }
            //运行时启动参数

            if(ExportExcel.Export(txtbox.Text,luatxtbox.Text, translateType.ToArray(),txtbox_filter.Text)){
                ShowTips("导出完成");
            }
        }

        /// <summary>
        /// 删除translate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btn_translate_Del(object sender, EventArgs e)
        {
            TextBox txtbox = this.Controls[7] as TextBox;
            if (ExportExcel.DelTranslate(txtbox.Text))
            {
                ShowTips("删除完成");
            }
        }

        /// <summary>
        /// 选择删除translate目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btn_del_trans_Click(object sender, EventArgs e) {
            OpenFileDialog openfile = new OpenFileDialog();
            DialogResult result = openfile.ShowDialog();
            if (result == DialogResult.OK)
            {
                TextBox txtbox = this.Controls[7] as TextBox;
                txtbox.Text = openfile.FileName;
                int index = txtbox.Text.LastIndexOf(@"\");
                txtbox.Text = txtbox.Text.Substring(0, index);
            }
        }



        /// <summary>
        /// 选择过滤目录路劲
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btn_filter_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            DialogResult result = openfile.ShowDialog();
            if (result == DialogResult.OK)
            {
                TextBox txtbox = this.Controls[10] as TextBox;
                txtbox.Text = openfile.FileName;
                int index = txtbox.Text.LastIndexOf(@"\");
                txtbox.Text = txtbox.Text.Substring(0, index);
            }
        }

        /// <summary>
        /// 选择lua表翻译字段名数据文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btn_lua_Click(object sender, EventArgs e) {
            OpenFileDialog openfile = new OpenFileDialog();
            DialogResult result = openfile.ShowDialog();
            if (result == DialogResult.OK)
            {
                TextBox txtbox = this.Controls[16] as TextBox;
                txtbox.Text = openfile.FileName;
                int index = txtbox.Text.LastIndexOf(@"\");
                txtbox.Text = txtbox.Text.Substring(0, index);
            }
        }

        public void CmdRun(string[] args)
        {
            if (args.Length < 2) {
                ShowTips("必要参数至少要2个(rootpath,luaexportpath)");
                return;
            }

            List<string> translateType = new List<string>();
            string url_lan_xml = System.Environment.CurrentDirectory + @"\language.xml";
            //string[] urlPath = new string[4] { null, null, null, null};
            if (!parseXml(ref translateType, url_lan_xml)) {
                ShowTips("加载"+url_lan_xml+"语言配置文件出错");
                return;
            }

            if (args.Length >= 3)
            {
                if (ExportExcel.Export(args[0], args[1], translateType.ToArray(),args[2]))
                {
                    ShowTips("导出完成");
                }
            }
            else {
                if (ExportExcel.Export(args[0], args[1], translateType.ToArray()))
                {
                    ShowTips("导出完成");
                }
            }

            
        }

        public static void ShowTips(string content) {
            MessageBox.Show(
                    content,"提示消息",MessageBoxButtons.OK
                );
        }
    }
}
