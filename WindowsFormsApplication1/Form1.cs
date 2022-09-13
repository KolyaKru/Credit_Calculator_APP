using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double f1 = Convert.ToDouble(textBox1.Text);
            double f2 = Convert.ToDouble(textBox2.Text);
            double f3 = Convert.ToDouble(textBox3.Text);
            double k1 = Math.Floor(f3 / (12 * f1 - 1));    
            int m = 5;
            int n = Convert.ToInt32(f1 * 12 + 6);
            double k = 1;
            double sum1 = 0;
            double sum2 = 0;
            double sum3 = 0;
            string[,] a = new string[n, m];
            a[5, 0] = "Номер месяца";
            a[5, 1] = "Остаток кредита,$";
            a[5, 2] = "Месячный долг,$";
            a[5, 3] = "Сумма процентов,$";
            a[5, 4] = "Общий платёж,$";
            a[0, 0] = "Кредит пользователя";
            a[1, 0] = "Срок кредита, год";
            a[2, 0] = "Процентная ставка, %";
            a[3, 0] = "Сумма кредита, $";
            a[1, 1] = textBox1.Text;
            a[2, 1] = textBox2.Text;
            a[3, 1] = textBox3.Text;
            for (int i = 6; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    if (j == 0)
                    {
                        if (i == n - 1)
                            a[i, j] = "ИТОГО:";
                        else
                        {
                            a[i, j] = Convert.ToString(k);

                            if (k == 12)
                                k = 0;
                            k++;
                        }


                    }
                    else
                        if (j == 1)
                        {
                            if (i == n - 1)
                                a[i, j] = "0";
                            else
                            {
                                if (i == 6)
                                {
                                    a[i, j] = Convert.ToString(f3);

                                }
                                else
                                    a[i, j] = Convert.ToString(Convert.ToDouble(a[i - 1, j]) - Math.Floor(f3 / (12 * f1 - 1)));
                            }
                        }
                        else
                            if (j == 2)
                            {
                                if (i == n - 1)
                                    a[i, j] = Convert.ToString(sum1);
                                else
                                {
                                    a[i, j] = Convert.ToString(Math.Floor(f3 / (12 * f1 - 1)));

                                    if (i == n - 2)
                                        a[i, j] = Convert.ToString(a[i, j - 1]);
                                    sum1 = sum1 + Convert.ToDouble(a[i, j]);
                                }

                            }
                            else
                                if (j == 3)
                                {
                                    if (i == n - 1)
                                        a[i, j] = Convert.ToString(sum2);
                                    else
                                    {
                                        a[i, j] = Convert.ToString(Math.Floor((Convert.ToDouble(a[i, 1]) * 30 * f2) / 360 / 100));
                                        sum2 = sum2 + Convert.ToDouble(a[i, j]);
                                    }
                                }
                                else
                                    if (j == 4)
                                    {
                                        if (i == n - 1)
                                            a[i, j] = Convert.ToString(sum3);
                                        else
                                        {
                                            a[i, j] = Convert.ToString(Convert.ToDouble(a[i, 2]) + Convert.ToDouble(a[i, 3]));
                                            sum3 = sum3 + Convert.ToDouble(a[i, j]);
                                        }

                                    }

                }
            }
            dataGridView1.RowCount = n;
            dataGridView1.ColumnCount = m;
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = a[i, j];
                }
            }
            textBox4.Text = Convert.ToString(sum3);
            textBox5.Text = Convert.ToString(sum2);
            textBox6.Text = Convert.ToString(k1);
            
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void сохранитьToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            SaveTable(dataGridView1);
            
        }

        void SaveTable(DataGridView What_save)
        {
            string path = System.IO.Directory.GetCurrentDirectory() + @"\" + "Отчёт по кредиту пользователя.xlsx";

            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            for (int i = 1; i < dataGridView1.RowCount + 1; i++)
            {
                for (int j = 1; j < dataGridView1.ColumnCount + 1; j++)
                {
                    worksheet.Rows[i].Columns[j] = dataGridView1.Rows[i - 1].Cells[j - 1].Value;
                }
            }
            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            excelapp.Quit();
        }
            
    }
}
