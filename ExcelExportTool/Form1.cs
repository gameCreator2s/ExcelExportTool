using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            this.Size = new Size(scr.Bounds.Width/2,scr.Bounds.Height/2);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Excel翻译表导出";
            //0
            Label label1 = new Label();
            label1.Location=new Point(24,8);
            label1.Size = new Size(500, 20);
            label1.Text="请选择要用于生成翻译表的原excel表所在目录";
            this.Controls.Add(label1);

            //1
            TextBox txtbox = new TextBox();
            txtbox.Location = new Point(24, 25);
            txtbox.Size = new Size(500, 20);
            txtbox.Text = @"E:\zzhx\trunk\data";
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
            string[] translatetype = new string[] { "ch", "en", "ru", "ls", "ch", "en", "ru", "ls" };
            //3
            GroupBox gbox = new GroupBox();
            gbox.Location = new Point(24, 100);
            gbox.Size = new Size(300, 200);

            Label label_bg = new Label();
            label_bg.Location = new Point(24, 30);
            label_bg.Size = new Size(200, 20);
            label_bg.Text = "选择要添加的语言种类";
            gbox.Controls.Add(label_bg);

            int preIndex = 0;
            int curIndex = 0;
            for (int i = 0; i < translatetype.Length; i++) {
                CheckBox cbox = new CheckBox();
                curIndex++;
                if (curIndex % 2 == 0)
                    preIndex++;
                cbox.Location = new Point(24+i%2*50, 50+preIndex*30);
                cbox.Size = new Size(200, 30);
                cbox.Text = translatetype[i];
                gbox.Controls.Add(cbox);
            }

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
            ExportExcel.Export(txtbox.Text, translateType.ToArray(), txtbox_filter.Text);
            ShowTips("导出完成");
        }

        /// <summary>
        /// 删除translate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btn_translate_Del(object sender, EventArgs e)
        {
            TextBox txtbox = this.Controls[7] as TextBox;
            ExportExcel.DelTranslate(txtbox.Text);
            ShowTips("删除完成");
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


        public static void ShowTips(string content) {
            MessageBox.Show(
                    content,"提示消息",MessageBoxButtons.OK
                );
        }
    }
}
