using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Spectre_of_Signal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            button7.Enabled = false;
            button2.Enabled = false;
            button6.Enabled = false;
            button3.Enabled = false;
        }
        public const int kznsize= 100000;
        public const int tznsize = 20002;
        public string[] kznstr;
        public string[,] tznstr;
        public double[] kzndbl;
        public double[,] tzndbl;
        public int  linesCount;
        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Created by EKNM");
        }
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string rfname = @"C:\r.txt";
            OpenFileDialog open = new OpenFileDialog();
            open.InitialDirectory = "С:\\";
            open.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            open.FilterIndex = 1;
            open.Title = "Открыть файл";
            if (open.ShowDialog() == DialogResult.OK)
            {
                rfname = open.FileName;
                using (var streamReader = new StreamReader(rfname, Encoding.UTF8))
                {
                    kznstr = new string[kznsize];
                    kznstr[0] = "k";
                    streamReader.ReadLine();
                    for (int r = 1; r < kznsize; r++)
                    {
                        string str = streamReader.ReadLine();
                        kznstr[r] = str.Substring(0, 14);                     
                    }
                    streamReader.Close();
                }
            }
            button1.Enabled = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string rfname23 = @"C:\t.txt";
            OpenFileDialog open = new OpenFileDialog();
            open.InitialDirectory = "С:\\";
            open.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            open.FilterIndex = 1;
            open.Title = "Открыть файл";
            if (open.ShowDialog() == DialogResult.OK)
            {
                rfname23 = open.FileName;
                using (var streamReaderrt = new StreamReader(rfname23, Encoding.UTF8))
                {
                    tznstr = new string[tznsize, 2];
                    tznstr[0,0] = "t";
                    tznstr[0, 1] = "f(t)";
                    for (int r = 0; r < tznsize-1; r++)
                    {
                        string str = streamReaderrt.ReadLine();
                        tznstr[r+1, 0] = str.Substring(0,14);
                        tznstr[r + 1, 1] = str.Substring(15, 15);
                    }
                    streamReaderrt.Close();
                }
            }
            button5.Enabled = false;
            button7.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            kzndbl = new double[kznsize];
            tzndbl = new double[tznsize,2];
            for (int r = 1; r < kznsize; r++)
            {
                double kfv = Convert.ToDouble(kznstr[r][0].ToString());
                double kdrob = Convert.ToDouble(kznstr[r].Substring(2, 8));
                kdrob /= 100000000;
                kfv += kdrob;
                double kvstp = Convert.ToInt32(kznstr[r][13].ToString());
                if(kznstr[r].IndexOf('-')==-1)
                {
                    kvstp = Math.Pow(10, kvstp);
                }
                else
                {
                    kvstp = Math.Pow(0.1, kvstp);
                }
                kzndbl[r] = kfv* kvstp;
            }
            for (int r = 1; r < tznsize; r++)
            {
                double tfv = Convert.ToDouble(tznstr[r,0][0].ToString());
                double ftfv = Convert.ToDouble(tznstr[r, 1][1].ToString());
                double tdrob= Convert.ToDouble(tznstr[r, 0].Substring(2, 8));
                tdrob /= 100000000;
                tfv += tdrob;
                double ftdrob = Convert.ToDouble(tznstr[r, 1].Substring(3, 8));
                ftdrob /= 100000000;
                ftfv += ftdrob;
                double tstp = Convert.ToInt32(tznstr[r,0][13].ToString());
                double ftstp = Convert.ToInt32(tznstr[r, 1][14].ToString());
                if (tznstr[r,0].IndexOf('-') == -1)
                {
                    tstp = Math.Pow(10, tstp);
                }
                else
                {
                    tstp = Math.Pow(0.1, tstp);
                }
                if (tznstr[r, 1][0] == '-')
                {
                    ftfv *= -1;
                }
                if (tznstr[r, 1][11] == '-')
                {
                    ftstp = Math.Pow(0.1, ftstp);
                }
                else
                {
                    ftstp = Math.Pow(10, ftstp);
                }
                tzndbl[r,0] = tfv*tstp;
                tzndbl[r, 1] =ftfv*ftstp;
            }
            dataGridView1.ColumnCount = 2;
            dataGridView1.Rows[0].Cells[1].Value = kznstr[0] + " conv";
            for (int r = 1; r < kznsize; r++)
            {
                dataGridView1.Rows[r].Cells[1].Value = kzndbl[r].ToString();
            }
            dataGridView2.ColumnCount = 4;
            dataGridView2.Rows[0].Cells[2].Value = tznstr[0, 0] + " conv";
            dataGridView2.Rows[0].Cells[3].Value = tznstr[0, 1] + " conv";
            for (int r = 1; r < tznsize; r++)
            {
                dataGridView2.Rows[r].Cells[2].Value = tzndbl[r, 0].ToString();
                dataGridView2.Rows[r].Cells[3].Value = tzndbl[r, 1].ToString();
            }
            button2.Enabled = false;
            button6.Enabled = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = kznsize;
            dataGridView1.ColumnCount = 1;
            dataGridView1.Rows[0].Cells[0].Value = kznstr[0];
            for (int r = 1; r < kznsize; r++)
            {
                dataGridView1.Rows[r].Cells[0].Value = kznstr[r];
            }
            dataGridView2.RowCount = tznsize;
            dataGridView2.ColumnCount = 2;
            dataGridView2.Rows[0].Cells[0].Value = tznstr[0, 0];
            dataGridView2.Rows[0].Cells[1].Value = tznstr[0, 1];
            for (int r = 1; r < tznsize; r++)
            {
                dataGridView2.Rows[r].Cells[0].Value = tznstr[r, 0];
                dataGridView2.Rows[r].Cells[1].Value = tznstr[r, 1];
            }
            button7.Enabled = false;
            button2.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheetk = (Microsoft.Office.Interop.Excel._Worksheet)workbook.Worksheets.Add(); 
            app.Visible = true;
            worksheetk = workbook.ActiveSheet;
            worksheetk.Name = "k";
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheetk.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            } 
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheetk.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            Microsoft.Office.Interop.Excel._Worksheet worksheett = (Microsoft.Office.Interop.Excel._Worksheet)workbook.Worksheets.Add();
            app.Visible = true;
            worksheett = workbook.ActiveSheet;
            worksheett.Name = "t";  
            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
            {
                worksheett.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
            } 
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    worksheett.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                }
            }
            workbook.SaveAs("output.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Quit();
            button1.Enabled = true;
            button5.Enabled = true;
            button3.Enabled = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            button6.Enabled = false;
            button3.Enabled = true;
        }
    }
}
