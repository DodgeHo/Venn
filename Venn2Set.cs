using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace Venn
{
    public partial class Venn2Set : Form
    {
        public Venn2Set(List<string> Names, List<string> Texts)
        {
            
            InitializeComponent();
            this.Names = Names;
            this.Texts = Texts;
        }

        private void Venn2Set_Load(object sender, EventArgs e)
        {


            DataTable dt = new DataTable();//建立个数据表
            dt.Columns.Add(new DataColumn("Set Name", typeof(string)));//在表中添加string类型的Name列
            dt.Columns.Add(new DataColumn("nitems", typeof(int)));//在表中添加int类型的列
            dt.Columns.Add(new DataColumn("Element", typeof(string)));//在表中添加string类型的Name列

            HashSet<string> SetA = GetElement(Texts[0]);
            HashSet<string> SetB = GetElement(Texts[1]);

            HashSet<string> pureA_B=SetA.ToHashSet<string>(); pureA_B.IntersectWith(SetB);
            HashSet<string> pureA = SetA.ToHashSet<string>(); pureA.ExceptWith(pureA_B);
            HashSet<string> pureB = SetB.ToHashSet<string>(); pureB.ExceptWith(pureA_B);
            HashSet<string> Total = SetA.ToHashSet<string>(); Total.UnionWith(SetB);

            DataRow dr;//行

            dr = dt.NewRow();
            dr["Set Name"] = Names[0]+" & "+Names[1];
            dr["nitems"] = pureA_B.Count;
            dr["Element"] = ElementToString(pureA_B);
            dt.Rows.Add(dr);//在表的对象的行里添加此行

            dr = dt.NewRow();
            dr["Set Name"] = Names[0] ;
            dr["nitems"] = pureA.Count;
            dr["Element"] = ElementToString(pureA);
            dt.Rows.Add(dr);//在表的对象的行里添加此行

            dr = dt.NewRow();
            dr["Set Name"] =Names[1];
            dr["nitems"] = pureB.Count;
            dr["Element"] = ElementToString(pureB);
            dt.Rows.Add(dr);//在表的对象的行里添加此行

            dr = dt.NewRow();
            dr["Set Name"] = "Total";
            dr["nitems"] = Total.Count;
            dr["Element"] = ElementToString(Total);
            dt.Rows.Add(dr);//在表的对象的行里添加此行

            dataGridView1.DataSource = dt;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

            pictureBox1.Controls.Add(label1);
            label1.Location = new Point(88, 166);
            pictureBox1.Controls.Add(label2);
            label2.Location = new Point(192, 166);
            pictureBox1.Controls.Add(label3);
            label3.Location = new Point(320, 166);
            pictureBox1.Controls.Add(label4);
            label4.Location = new Point(28, 326);
            pictureBox1.Controls.Add(label5);
            label5.Location = new Point(28, 346);


            label1.Text = pureA.Count.ToString();
            label2.Text = pureA_B.Count.ToString();
            label3.Text = pureB.Count.ToString();
            label4.Text = Names[0].ToString();
            label5.Text = Names[1].ToString();
        }





        List<string> Names;
        List<string> Texts;

        public void DrawGroundTruth()
        {
            Graphics g = this.CreateGraphics();
            //g.Clear(this.BackColor);
            //DrawPerson();
            //画一幅图像 
            Image curImage = Venn.Properties.Resources.button_over;
            g.DrawImage(curImage, 400, 400, curImage.Width, curImage.Height);
            g.Dispose();
        }



        public HashSet<string> GetElement(string Text)
        {
            
            
            HashSet<string> output = new HashSet<string> ();
            output.Clear();
            string[] ss = Text.Split('\n');

            foreach (string text in ss)
            {
                string subtext = text.Trim();
                if (subtext.Length > 0)
                    output.Add(subtext);
            }
            return output;
        }

        public string ElementToString(HashSet<string> input)
        {
            if (input.Count == 0)
                return "";
            string output = "";

            foreach (String st in input)
            {
                // output += st + "\n\r";
                output += st + ", ";
            }
            output = output.Substring(0, output.Length - 2);
            return output;
        }
        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Button5_Click(object sender, EventArgs e)
        {
            ExportToExcel d = new ExportToExcel();
            d.OutputAsExcelFile(dataGridView1);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Bitmap bit = new Bitmap(pictureBox1.Width, pictureBox1.Height);//实例化一个和窗体一样大的bitmap
            Graphics g = Graphics.FromImage(bit);
            g.CompositingQuality = CompositingQuality.HighQuality;//质量设为最高
            //g.CopyFromScreen(this.Left+pictureBox1.Left, this.Top+pictureBox1.Top, 0, 0, new Size(pictureBox1.Width, pictureBox1.Height));//保存整个窗体为图片
            g.CopyFromScreen(pictureBox1.PointToScreen(Point.Empty), Point.Empty, pictureBox1.Size);
            SaveToPNG p = new SaveToPNG();
            p.OutputAsPNGFile(bit);
        }

    }


}


