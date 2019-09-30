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

using System.Collections;

namespace Venn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ElementButtonType = 3; SetVisible();
            button2.BackgroundImage = Venn.Properties.Resources.button_on;
            textBox2.BackgroundImage = textBox3.BackgroundImage= textBox5.BackgroundImage
                = textBox7.BackgroundImage= textBox9.BackgroundImage= 
                Venn.Properties.Resources.blank;
            textBox1.Text = "set A";
            textBox4.Text = "set B";
            textBox6.Text = "set C";
            textBox8.Text = "set D";
            textBox10.Text = "set E";


        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void CheckedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            ResetButtonBackground();
            ElementButtonType = 2;
            button1.BackgroundImage = Venn.Properties.Resources.button_on;
            SetVisible();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            ResetButtonBackground();
            ElementButtonType = 3;
            button2.BackgroundImage = Venn.Properties.Resources.button_on;
            SetVisible();
        }

        private void ResetButtonBackground() {
            button1.BackgroundImage = Venn.Properties.Resources.button_off;
            button2.BackgroundImage = Venn.Properties.Resources.button_off;
            button3.BackgroundImage = Venn.Properties.Resources.button_off;
            button4.BackgroundImage = Venn.Properties.Resources.button_off;
        }


        static int ElementButtonType;

        private void Button3_Click(object sender, EventArgs e)
        {
            ResetButtonBackground();
            ElementButtonType = 4;
            button3.BackgroundImage = Venn.Properties.Resources.button_on;
            SetVisible();
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            ResetButtonBackground();
            ElementButtonType = 5;
            button4.BackgroundImage = Venn.Properties.Resources.button_on;
            SetVisible();
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void FlowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
            
            e.Graphics.Clear(flowLayoutPanel1.BackColor);
            e.Graphics.DrawString("Set 1 ", flowLayoutPanel1.Font, Brushes.DodgerBlue, 18, 18);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, e.Graphics.MeasureString(flowLayoutPanel1.Text, flowLayoutPanel1.Font).Width + 8, 7, flowLayoutPanel1.Width - 2, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 1, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, flowLayoutPanel1.Height - 2, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, flowLayoutPanel1.Width - 2, 7, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
        }


        private void FlowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(flowLayoutPanel1.BackColor);
            e.Graphics.DrawString("Set 2 ", flowLayoutPanel1.Font, Brushes.DarkRed, 18, 18);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, e.Graphics.MeasureString(flowLayoutPanel1.Text, flowLayoutPanel1.Font).Width + 8, 7, flowLayoutPanel1.Width - 2, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 1, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, flowLayoutPanel1.Height - 2, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, flowLayoutPanel1.Width - 2, 7, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
        }

        private void FlowLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(flowLayoutPanel1.BackColor);
            e.Graphics.DrawString("Set 3 ", flowLayoutPanel1.Font, Brushes.ForestGreen, 18, 18);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, e.Graphics.MeasureString(flowLayoutPanel1.Text, flowLayoutPanel1.Font).Width + 8, 7, flowLayoutPanel1.Width - 2, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 1, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, flowLayoutPanel1.Height - 2, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, flowLayoutPanel1.Width - 2, 7, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
        }

        private void FlowLayoutPanel4_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(flowLayoutPanel1.BackColor);
            e.Graphics.DrawString("Set 4 ", flowLayoutPanel1.Font, Brushes.RosyBrown, 18, 18);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, e.Graphics.MeasureString(flowLayoutPanel1.Text, flowLayoutPanel1.Font).Width + 8, 7, flowLayoutPanel1.Width - 2, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 1, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, flowLayoutPanel1.Height - 2, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, flowLayoutPanel1.Width - 2, 7, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
        }

        private void FlowLayoutPanel5_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(flowLayoutPanel1.BackColor);
            e.Graphics.DrawString("Set 5 ", flowLayoutPanel1.Font, Brushes.MediumPurple, 18, 18);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, e.Graphics.MeasureString(flowLayoutPanel1.Text, flowLayoutPanel1.Font).Width + 8, 7, flowLayoutPanel1.Width - 2, 7);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, 7, 1, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, 1, flowLayoutPanel1.Height - 2, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
            e.Graphics.DrawLine(Pens.DodgerBlue, flowLayoutPanel1.Width - 2, 7, flowLayoutPanel1.Width - 2, flowLayoutPanel1.Height - 2);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            button5.BackgroundImage = Venn.Properties.Resources.button_on;
            List<string> Names =new List<string>();
            List<string> Texts =new List<string>();

            Names.Clear();Texts.Clear();
            Names.Add(textBox1.Text);
            Names.Add(textBox4.Text);
            Texts.Add(textBox2.Text);
            Texts.Add(textBox3.Text);
            switch (ElementButtonType)
            {
                case 2:
                    
                    Venn2Set f2 = new Venn2Set(Names,Texts);
                    f2.Show();
                    break;

                case 3:
                    Names.Add(textBox6.Text);
                    Texts.Add(textBox5.Text);
                    Venn3Set f3 = new Venn3Set(Names, Texts);
                    f3.Show();
                    break;

                case 4:
                    Names.Add(textBox6.Text);
                    Texts.Add(textBox5.Text);
                    Names.Add(textBox8.Text);
                    Texts.Add(textBox7.Text);
                    Venn4Set f4 = new Venn4Set(Names, Texts);
                    f4.Show();
                    break;

                case 5:
                    Names.Add(textBox6.Text);
                    Texts.Add(textBox5.Text);
                    Names.Add(textBox8.Text);
                    Texts.Add(textBox7.Text);
                    Names.Add(textBox10.Text);
                    Texts.Add(textBox9.Text);
                    Venn5Set f5 = new Venn5Set(Names, Texts);
                    f5.Show();
                    break;
            }
            button5.BackgroundImage = Venn.Properties.Resources.button_off;
            Names.Clear(); Texts.Clear();
        }

        private void SetVisible() {
            flowLayoutPanel3.Visible = false;
            flowLayoutPanel4.Visible = false;
            flowLayoutPanel5.Visible = false;

            if (ElementButtonType>2)
                flowLayoutPanel3.Visible = true;
            if (ElementButtonType > 3)
                flowLayoutPanel4.Visible = true;
            if (ElementButtonType > 4)
                flowLayoutPanel5.Visible = true;
        }

        private void Label11_Click(object sender, EventArgs e)
        {

        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            
            textBox1.Text = "set A";
            textBox4.Text = "set B";
            textBox6.Text = "set C";
            textBox8.Text = "set D";
            textBox10.Text = "set E";

            textBox2.Text=" ";
            textBox3.Text=" ";
            textBox7.Text=" ";
            textBox9.Text=" ";
            textBox5.Text=" ";
        }

        private void FlowLayoutPanel6_Paint(object sender, PaintEventArgs e)
        {

        }
    }



    public class SaveToPNG
    {
        public void OutputAsPNGFile(Bitmap bit)
        {

            string filePath = "";
            SaveFileDialog s = new SaveFileDialog();
            s.Title = "Save PNG picture";
            s.Filter = "PNG picture(*.png)|*.png";
            s.FilterIndex = 1;
            if (s.ShowDialog() == DialogResult.OK)
                filePath = s.FileName;
            else
                return;

            bit.Save(filePath);//默认保存格式为PNG，保存成jpg格式质量不是很好
            MessageBox.Show("Export Succeed！", "CAPION", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    public class ExportToExcel
    {
        public Excel.Application m_xlApp = null;

        public void OutputAsExcelFile(DataGridView dataGridView)
        {
            if (dataGridView.Rows.Count <= 0)
            {
                MessageBox.Show("No Data！", "CAPION", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            string filePath = "";
            SaveFileDialog s = new SaveFileDialog();
            s.Title = "Save Excel file";
            s.Filter = "Excel file(*.xls)|*.xls";
            s.FilterIndex = 1;
            if (s.ShowDialog() == DialogResult.OK)
                filePath = s.FileName;
            else
                return;

            //第一步：将dataGridView转化为dataTable,这样可以过滤掉dataGridView中的隐藏列  

            DataTable tmpDataTable = new DataTable("tmpDataTable");
            DataTable modelTable = new DataTable("ModelTable");
            for (int column = 0; column < dataGridView.Columns.Count; column++)
            {
                if (dataGridView.Columns[column].Visible == true)
                {
                    DataColumn tempColumn = new DataColumn(dataGridView.Columns[column].HeaderText, typeof(string));
                    tmpDataTable.Columns.Add(tempColumn);
                    DataColumn modelColumn = new DataColumn(dataGridView.Columns[column].Name, typeof(string));
                    modelTable.Columns.Add(modelColumn);
                }
            }
            for (int row = 0; row < dataGridView.Rows.Count; row++)
            {
                if (dataGridView.Rows[row].Visible == false)
                    continue;
                DataRow tempRow = tmpDataTable.NewRow();
                for (int i = 0; i < tmpDataTable.Columns.Count; i++)
                    tempRow[i] = dataGridView.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value;
                tmpDataTable.Rows.Add(tempRow);
            }
            if (tmpDataTable == null)
            {
                return;
            }

            //第二步：导出dataTable到Excel  
            long rowNum = tmpDataTable.Rows.Count;//行数  
            int columnNum = tmpDataTable.Columns.Count;//列数  
            Excel.Application m_xlApp = new Excel.Application();
            m_xlApp.DisplayAlerts = false;//不显示更改提示  
            m_xlApp.Visible = false;

            Excel.Workbooks workbooks = m_xlApp.Workbooks;
            Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];//取得sheet1  

            try
            {
                string[,] datas = new string[rowNum + 1, columnNum];
                for (int i = 0; i < columnNum; i++) //写入字段  
                    datas[0, i] = tmpDataTable.Columns[i].Caption;
                //Excel.Range range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]);  
                Excel.Range range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]];
                range.Interior.ColorIndex = 15;//15代表灰色  
                range.Font.Bold = true;
                range.Font.Size = 10;

                int r = 0;
                for (r = 0; r < rowNum; r++)
                {
                    for (int i = 0; i < columnNum; i++)
                    {
                        object obj = tmpDataTable.Rows[r][tmpDataTable.Columns[i].ToString()];
                        datas[r + 1, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式  
                    }
                    System.Windows.Forms.Application.DoEvents();
                    //添加进度条  
                }
                //Excel.Range fchR = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]);  
                Excel.Range fchR = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];
                fchR.Value2 = datas;

                worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。  
                                                         //worksheet.Name = "dd";  

                //m_xlApp.WindowState = Excel.XlWindowState.xlMaximized;  
                m_xlApp.Visible = false;

                // = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]);  
                range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];

                //range.Interior.ColorIndex = 15;//15代表灰色  
                range.Font.Size = 9;
                range.RowHeight = 14.25;
                range.Borders.LineStyle = 1;
                range.HorizontalAlignment = 1;
                workbook.Saved = true;
                workbook.SaveCopyAs(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export Error：" + ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                EndReport();
            }

            m_xlApp.Workbooks.Close();
            m_xlApp.Workbooks.Application.Quit();
            m_xlApp.Application.Quit();
            m_xlApp.Quit();
            MessageBox.Show("Export Succeed！", "CAPION", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void EndReport()
        {
            object missing = System.Reflection.Missing.Value;
            try
            {
                //m_xlApp.Workbooks.Close();  
                //m_xlApp.Workbooks.Application.Quit();  
                //m_xlApp.Application.Quit();  
                //m_xlApp.Quit();  
            }
            catch { }
            finally
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Application);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp);
                    m_xlApp = null;
                }
                catch { }
                try
                {
                    //清理垃圾进程  
                    this.killProcessThread();
                }
                catch { }
                GC.Collect();
            }
        }

        private void killProcessThread()
        {
            ArrayList myProcess = new ArrayList();
            for (int i = 0; i < myProcess.Count; i++)
            {
                try
                {
                    System.Diagnostics.Process.GetProcessById(int.Parse((string)myProcess[i])).Kill();
                }
                catch { }
            }
        }
    }

}
