using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Drawing.Drawing2D;
using Excel = Microsoft.Office.Interop.Excel;
namespace Venn
{
    public partial class Venn3Set : Form
    {


        public Venn3Set(List<string> Names, List<string> Texts)
        {

            InitializeComponent();
            this.Names = Names;
            this.Texts = Texts;
        }

        private void Venn3Set_Load(object sender, EventArgs e)
        {


            DataTable dt = new DataTable();//建立个数据表
            dt.Columns.Add(new DataColumn("Set Name", typeof(string)));//在表中添加string类型的Name列
            dt.Columns.Add(new DataColumn("nitems", typeof(int)));//在表中添加int类型的列
            dt.Columns.Add(new DataColumn("Element", typeof(string)));//在表中添加string类型的Name列

            HashSet<string> SetA = GetElement(Texts[0]);
            HashSet<string> SetB = GetElement(Texts[1]);
            HashSet<string> SetC = GetElement(Texts[2]);

            HashSet<string> pureA_B_C = SetA.ToHashSet<string>(); pureA_B_C.IntersectWith(SetB); pureA_B_C.IntersectWith(SetC);
            HashSet<string> Total = SetA.ToHashSet<string>(); Total.UnionWith(SetB);Total.UnionWith(SetC);
            HashSet<string> pureA_B = SetA.ToHashSet<string>(); pureA_B.IntersectWith(SetB); pureA_B.ExceptWith(SetC);
            HashSet<string> pureA_C = SetA.ToHashSet<string>(); pureA_C.IntersectWith(SetC); pureA_C.ExceptWith(SetB);
            HashSet<string> pureB_C = SetB.ToHashSet<string>(); pureB_C.IntersectWith(SetC); pureB_C.ExceptWith(SetA);
            HashSet<string> pureA = SetA.ToHashSet<string>(); pureA.ExceptWith(SetB); pureA.ExceptWith(SetC);
            HashSet<string> pureB = SetB.ToHashSet<string>(); pureB.ExceptWith(SetA); pureB.ExceptWith(SetC);
            HashSet<string> pureC = SetC.ToHashSet<string>(); pureC.ExceptWith(SetA); pureC.ExceptWith(SetB);

            HashSet<string>[] Setgroup = {pureA, pureB, pureC, 
                                    pureA_B,pureA_C,pureB_C,
                                    pureA_B_C};
            string[] Namegroup = { Names[0], Names[1], Names[2],
                            Names[0] + " & " + Names[1], Names[0] + " & " + Names[2],
                            Names[1] + " & " + Names[2],
                            Names[0] + " & " + Names[1] + " & " + Names[2] };





            DataRow dr;//行
            
            for (int i = Setgroup.Length-1; i>=0; i--)
            {
                dr = dt.NewRow();
                dr["Set Name"] = Namegroup[i];
                dr["nitems"] = Setgroup[i].Count;
                dr["Element"] = ElementToString(Setgroup[i]);
                dt.Rows.Add(dr);//在表的对象的行里添加此行
            }

            dr = dt.NewRow();
            dr["Set Name"] = "Total";
            dr["nitems"] = Total.Count;
            dr["Element"] = ElementToString(Total);
            dt.Rows.Add(dr);//在表的对象的行里添加此行


            dataGridView1.DataSource = dt;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

            Label[] labs = { label1, label2,label3,An,Bn,Cn,ABn,BCn,ACn,ABCn};
            foreach(Label et in labs){
                pictureBox1.Controls.Add(et);
                et.Location = new Point(et.Left - pictureBox1.Left, et.Top - pictureBox1.Top);
            }
            Cn.Text = pureC.Count.ToString();
            An.Text = pureA.Count.ToString();
            Bn.Text = pureB.Count.ToString();
            ACn.Text = pureA_C.Count.ToString();
            ABn.Text = pureA_B.Count.ToString();
            BCn.Text = pureB_C.Count.ToString();
            ABCn.Text = pureA_B_C.Count.ToString();

            label1.Text = Names[0].ToString();
            label2.Text = Names[1].ToString();
            label3.Text = Names[2].ToString();
        }





        List<string> Names;
        List<string> Texts;



        public HashSet<string> GetElement(string Text)
        {


            HashSet<string> output = new HashSet<string>();
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

        private void BCn_Click(object sender, EventArgs e)
        {

        }

        private void ACn_Click(object sender, EventArgs e)
        {

        }

        private void Bn_Click(object sender, EventArgs e)
        {

        }

        private void An_Click(object sender, EventArgs e)
        {

        }

        private void Cn_Click(object sender, EventArgs e)
        {

        }

        private void DataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ABCn_Click(object sender, EventArgs e)
        {

        }

        private void ABn_Click(object sender, EventArgs e)
        {

        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }


        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }
    }




}
