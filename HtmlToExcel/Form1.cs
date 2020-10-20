using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HtmlToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 選擇Excel腳本並讀取Sheet到ComboBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scriptbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Select Script File"
            };
            string path = Directory.GetCurrentDirectory() + @"\Script";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            dialog.InitialDirectory = Directory.GetCurrentDirectory() + @"\Script";
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                scripttextBox.Text = dialog.FileName;
                FileInfo file = new FileInfo(dialog.FileName);
                comboBox1.Items.Clear();
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    foreach(ExcelWorksheet excelWorksheet in excelPackage.Workbook.Worksheets)
                    {
                        comboBox1.Items.Add(excelWorksheet.Name);
                    }
                    if(comboBox1.Items.Count>0) //有工作表就顯示
                        comboBox1.SelectedItem = comboBox1.Items[0];
                }
            }
        }

        /// <summary>
        /// 點2下打開Script
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scripttextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (File.Exists(scripttextBox.Text))
            {
                System.Diagnostics.Process.Start(scripttextBox.Text);
            }
            else
            {
                MessageBox.Show("找不到Script\n"+scripttextBox.Text);
            }
        }

        /// <summary>
        /// 選擇要寫入的Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void excelbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Select Script File"
            };
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                exceltextBox.Text = dialog.FileName;
            }
        }

        /// <summary>
        /// 點2下開啟要寫入的Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exceltextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (File.Exists(exceltextBox.Text))
            {
                System.Diagnostics.Process.Start(exceltextBox.Text);
            }
            else
            {
                MessageBox.Show("找不到Excel\n" + exceltextBox.Text);
            }
        }

        /// <summary>
        /// 選擇要讀取的HTML
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void hmtlbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Html File";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Filter = "HTML|*.html";
            if (dialog.ShowDialog() == DialogResult.OK)
                htmltextBox.Text = dialog.FileName;
        }

        /// <summary>
        /// 點2下打開HTML
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void htmltextBox_DoubleClick(object sender, EventArgs e)
        {
            if (File.Exists(htmltextBox.Text))
            {
                System.Diagnostics.Process.Start(htmltextBox.Text);
            }
            else
            {
                MessageBox.Show("找不到HTML\n" + htmltextBox.Text);
            }
        }

        /// <summary>
        /// 限定輸入數字
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }

        /// <summary>
        /// 點Run執行寫入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void runbutton_Click(object sender, EventArgs e)
        {
            if (!File.Exists(scripttextBox.Text))
            {
                MessageBox.Show("找不到Script\n" + htmltextBox.Text);
                return;
            }
            if (!File.Exists(exceltextBox.Text))
            {
                MessageBox.Show("找不到Excel\n" + exceltextBox.Text);
                return;
            }
            if (!File.Exists(htmltextBox.Text))
            {
                MessageBox.Show("找不到HTML\n" + htmltextBox.Text);
                return;
            }
            if(htmlcolumntextBox.Text==""|| int.Parse(htmlcolumntextBox.Text)<0)
            {
                MessageBox.Show("請輸入正確Html Column");
                return;
            }
            //先檢查再寫入


        }

        /// <summary>
        /// 開始讀檔寫入
        /// </summary>
        /// <param name="script"></param>
        /// <param name="excel"></param>
        /// <param name="html"></param>
        private void WriteToExcel(FileInfo script,FileInfo excel ,FileInfo html)
        {

        }

    }
}
