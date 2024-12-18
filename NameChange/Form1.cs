using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace NameChange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string cellValue = null;
            bool stop = false;
            string textLine = "";
            string text1 = "";
            string finalText = "";
            do
            {
                string excelFilePath = string.Empty;
                openFileDialog1.InitialDirectory = Application.StartupPath;
                openFileDialog1.FileName = "*.xlsx";
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog1.FileName;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFilePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    cellValue = xlRange.Cells[1, 1].Value2.ToString();
                    textBox1.Text = cellValue;

                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    xlRange = null;
                    xlWorksheet = null;
                    xlWorkbook = null;
                    xlApp = null;
                    stop = true;

                    string textFilePath = string.Empty;
                    openFileDialog2.InitialDirectory = Application.StartupPath;
                    openFileDialog2.FileName = "*.txt";
                    openFileDialog2.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    if (openFileDialog2.ShowDialog() == DialogResult.OK)
                    {
                        textFilePath = openFileDialog2.FileName;
                        StreamReader sr = new StreamReader(textFilePath);
                        textLine = sr.ReadLine();
                        textBox2.Text = textLine;
                        text1 = textLine.Substring(0, 11);
                        sr.Close();
                        finalText = text1 + cellValue;
                        textBox3.Text = finalText;

                        StreamWriter sw = new StreamWriter(textFilePath);
                        sw.WriteLine(finalText);
                        sw.Close();
                        MessageBox.Show("İşlem tamamlandı!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //textBox1.Text = "";
                        //textBox2.Text = "";
                        //textBox3.Text = "";
                    }
                    else
                    {
                        MessageBox.Show("Dosya kaydedilemedi!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                    }
                }
                else
                {
                    //if (textBox1.Text == "")
                    //{
                        MessageBox.Show("Excel dosyasını seçmediniz!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    //}
                }
            }
            while (stop == false);
        }

    }
}
