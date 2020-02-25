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
        public string[,] datastr;
        public int  linesCount;
        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Created by EKNM");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // имя файла по умолчанию
            string rfname = @"C:\r.txt";
            OpenFileDialog open = new OpenFileDialog();
            // задание начальной директории (1)
            open.InitialDirectory = "С:\\";
            // задание свойства Filter (2)
            open.Filter = "srt files (*.txt)|*.txt|All files (*.*)|*.*";
            // задание свойства FilterIndex –исходный тип файла (3)
            open.FilterIndex = 1;
            // свойство Title - название окна диалога выбора файла (4)
            open.Title = "Открыть файл";
            // метод ShowDialog() показывает окно диалога
            if (open.ShowDialog() == DialogResult.OK)
            { // получаем имя файла для сохранения данных
                rfname = open.FileName; // выделяем память для потока - объекта типа
                using (var streamReader = new StreamReader(rfname, Encoding.UTF8))
                {
                    datastr = new string[20003, 3];
                    datastr[0, 0] = "k";
                    datastr[0, 1] = "t";
                    datastr[0, 2] = "rez";
                    for (int r = 6; r < 20002; r++)
                    {
                        string str = streamReader.ReadLine();
                        datastr[r, 0] = str.Substring(0, 14);                     
                    }
                    streamReader.Close();
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // имя файла по умолчанию
            string rfname23 = @"C:\t.txt";
            OpenFileDialog open = new OpenFileDialog();
            // задание начальной директории (1)
            open.InitialDirectory = "С:\\";
            // задание свойства Filter (2)
            open.Filter = "srt files (*.txt)|*.txt|All files (*.*)|*.*";
            // задание свойства FilterIndex –исходный тип файла (3)
            open.FilterIndex = 1;
            // свойство Title - название окна диалога выбора файла (4)
            open.Title = "Открыть файл";
            // метод ShowDialog() показывает окно диалога
            if (open.ShowDialog() == DialogResult.OK)
            { // получаем имя файла для сохранения данных
                rfname23 = open.FileName; // выделяем память для потока - объекта типа
                using (var streamReaderrt = new StreamReader(rfname23, Encoding.UTF8))
                {
                    for (int r = 0; r < 20001; r++)
                    {
                        string str = streamReaderrt.ReadLine();
                        datastr[r+1, 2] = str.Substring(0,14);
                    }
                    streamReaderrt.Close();
                }
            }
            button7.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = 20003;
            dataGridView1.ColumnCount = 3;
            dataGridView1.Rows[0].Cells[0].Value = datastr[0, 0];
            dataGridView1.Rows[0].Cells[1].Value = datastr[0, 1];
            dataGridView1.Rows[0].Cells[2].Value = datastr[0, 2];
            for (int r = 1; r < 20003; r++)
            {
                dataGridView1.Rows[r].Cells[0].Value = datastr[r, 0];
                dataGridView1.Rows[r].Cells[1].Value = datastr[r, 2];
                dataGridView1.Rows[r].Cells[2].Value = "Саня хуй саси";
            }
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[2].Width = 80;
        }
    }
}
