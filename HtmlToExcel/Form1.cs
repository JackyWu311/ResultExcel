using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace HtmlToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //建立script資料夾
            if (!Directory.Exists(Directory.GetCurrentDirectory() + @"\Script"))
            {
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\Script");
            }
            if (!Directory.Exists(Directory.GetCurrentDirectory() + @"\log"))
            {
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\log");
            }
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
                    foreach (ExcelWorksheet excelWorksheet in excelPackage.Workbook.Worksheets)
                    {
                        comboBox1.Items.Add(excelWorksheet.Name);
                    }
                    if (comboBox1.Items.Count > 0) //有工作表就顯示
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
                MessageBox.Show("找不到Script\n" + scripttextBox.Text);
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
            if (!CheckTextboxNull_or_NoSuchFile()) //先檢查textbox內容
                return;
            //紀錄開始時間
            DateTime time = DateTime.Now;
            //開啟Script excel
            using (ExcelPackage ScriptPackage = new ExcelPackage(new FileInfo(scripttextBox.Text)))
            {
                //要讀取的HTML欄位
                int Htmlcolumn = int.Parse(htmlcolumntextBox.Text);
                //指定ComboBox所選的Sheet
                ExcelWorksheet ScriptWorksheet = ScriptPackage.Workbook.Worksheets[comboBox1.Text];

                //script是空的或小於4行則return
                try
                {
                    if (ScriptWorksheet.Dimension.Rows < 4)
                    {
                        MessageBox.Show("請選擇正確格式工作表");
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("請選擇正確格式工作表");
                    return;
                }
                richTextBox1.Text = time.ToString("yyyyMMdd-HH:mm:ss") + "\nScript: " + scripttextBox.Text + "\nScript Sheet: " + ScriptWorksheet.Name + "\nExcel: " + exceltextBox.Text + "\nHTML: " + htmltextBox.Text + "\nHTML Column: " + Htmlcolumn + "\n\nStart Writing:\n------------------------------------------------------------------------------------------------------\n";

                //讀取指定的Html column並傳回dataTable
                DataTable dataTable = GetHtmlTable(htmltextBox.Text);

                //開啟目標Excel
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(exceltextBox.Text)))
                {
                    //要寫入的Excel的工作表清單
                    ExcelWorksheets excelWorksheets = excelPackage.Workbook.Worksheets;
                    //紀錄字串
                    List<string> HTMLSuccessString = new List<string> { "Success", "sucess", "Pass", "pass" };
                    List<string> HTMLFailString = new List<string> { "Fail", "fail", "Skip", "skip", "Stop", "stop" };
                    List<string> UsedSuccessString = new List<string>();  //紀錄寫入過的Success字眼
                    List<string> UsedFailString = new List<string>(); //紀錄寫入過的Fail字眼

                    bool MustFill = true;
                    for (int row = 4; row <= ScriptWorksheet.Dimension.Rows; row++)
                    {
                        //檢查Script必填欄位
                        MustFill = true;
                        string Mustfillstring = "[Not written]" + DateTime.Now.ToString(" yyyyMMdd-HH:mm:ss ") + ScriptPackage.File.Name + " Sheet: " + ScriptWorksheet.Name + " ";
                        if (ScriptWorksheet.Cells[row, 1].Value == null)  //htmlNo
                        {
                            Mustfillstring += ScriptWorksheet.Name + "[" + ScriptWorksheet.Cells[row, 1].ToString() + "]" + " HtmlNo. ，";
                            MustFill = false;
                        }
                        if (ScriptWorksheet.Cells[row, 2].Value == null) //Sheet
                        {
                            Mustfillstring += ScriptWorksheet.Name + "[" + ScriptWorksheet.Cells[row, 2].ToString() + "]" + " Sheet ，";
                            MustFill = false;
                        }
                        if (ScriptWorksheet.Cells[row, 3].Value == null) //Cell
                        {
                            Mustfillstring += ScriptWorksheet.Name + "[" + ScriptWorksheet.Cells[row, 3].ToString() + "]" + " Cell ，";
                            MustFill = false;
                        }
                        if (ScriptWorksheet.Cells[row, 4].Value == null) //Success
                        {
                            Mustfillstring += ScriptWorksheet.Name + "[" + ScriptWorksheet.Cells[row, 4].ToString() + "]" + " Success ，";
                            MustFill = false;
                        }
                        if (ScriptWorksheet.Cells[row, 6].Value == null) //Fail
                        {
                            Mustfillstring += ScriptWorksheet.Name + "[" + ScriptWorksheet.Cells[row, 6].ToString() + "]" + " Fail ，";
                            MustFill = false;
                        }
                        if (MustFill == false)
                        {
                            richTextBox1.AppendText(Mustfillstring);
                            continue;
                        }

                        //如果HtmlNo大於實際HTML Row則Continue不寫入
                        if (int.Parse(ScriptWorksheet.Cells[row, 1].Text) > dataTable.Rows.Count - 1)
                        {
                            richTextBox1.AppendText("[Not written]" + DateTime.Now.ToString(" yyyyMMdd-HH:mm:ss ") + ScriptPackage.File.Name + " Sheet: " + ScriptWorksheet.Name + "[" + ScriptWorksheet.Cells[row, 1].ToString() + "]" + " HtmlNo." + ScriptWorksheet.Cells[row, 1].Text + "大於實際HTML欄位" + (dataTable.Rows.Count - 1) + "\n");
                            continue;
                        }

                        //檢查目標excel是否含有script中的sheet
                        if (!excelWorksheets.Any(sheet => sheet.Name == ScriptWorksheet.Cells[row, 2].Text))
                        {
                            richTextBox1.AppendText("[Not written]" + DateTime.Now.ToString(" yyyyMMdd-HH:mm:ss ") + ScriptPackage.File.Name + " " + ScriptWorksheet.Name + "[" + ScriptWorksheet.Cells[row, 2].ToString() + "]" + " Sheet." + ScriptWorksheet.Cells[row, 2].Text + "不在 " + excelPackage.File.Name + " 中\n");
                            continue;
                        }

                        string WriteInLogString = "[Written]" + DateTime.Now.ToString(" yyyyMMdd-HH:mm:ss ") + ScriptPackage.File.Name + " Sheet: " + ScriptWorksheet.Name + " Row" + row + " \n             寫入 " + excelPackage.File.Name + " ";
                        //開始寫入
                        //如果讀到HTML Success
                        if (HTMLSuccessString.Contains((string)dataTable.Rows[int.Parse(ScriptWorksheet.Cells[row, 1].Text)][Htmlcolumn - 1]))
                        {
                            ExcelWorksheet excelWorksheet = excelWorksheets[ScriptWorksheet.Cells[row, 2].Text];
                            //如果是空欄位才寫入Success所有內容，否則跳過
                            if (excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Value == null)
                            {
                                //寫入Success以及附加內容
                                excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Value = ScriptWorksheet.Cells[row, 4].Text + ", " + ScriptWorksheet.Cells[row, 5].Text;
                                WriteInLogString += " Sheet: " + excelWorksheet.Name + " Cell[" + excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].ToString().ToUpper() + "]:\n             " + ScriptWorksheet.Cells[row, 4].Text + ", " + ScriptWorksheet.Cells[row, 5].Text + "\n";
                                //script註解欄有值再寫入註解
                                if (ScriptWorksheet.Cells[row, 8].Value != null)
                                {
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].AddComment(ScriptWorksheet.Cells[row, 8].Text, "User");
                                    WriteInLogString += "             新增註解: " + ScriptWorksheet.Cells[row, 8].Text + "\n";
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment.AutoFit = true;
                                }
                                //Note
                                if (ScriptWorksheet.Cells[row, 9].Value != null)
                                {
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].Value = ScriptWorksheet.Cells[row, 10].Text;
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].AutoFitColumns();
                                    WriteInLogString += "             Note[" + excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].ToString() + "]: " + ScriptWorksheet.Cells[row, 10].Text + "\n";
                                }
                            }
                            excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].AutoFitColumns();
                            UsedSuccessString.Add(ScriptWorksheet.Cells[row, 4].Text);
                            richTextBox1.AppendText(WriteInLogString);
                            continue;
                        }
                        //如果讀到HTML Fail
                        else if (HTMLFailString.Contains((string)dataTable.Rows[int.Parse(ScriptWorksheet.Cells[row, 1].Text)][Htmlcolumn - 1]))
                        {
                            ExcelWorksheet excelWorksheet = excelWorksheets[ScriptWorksheet.Cells[row, 2].Text];
                            //空欄位直接寫入 或 判斷欄位第一個字眼是否出現在success字眼，有則覆蓋
                            if (excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Value == null || UsedSuccessString.Contains(excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Text.Split(',')[0]))
                            {
                                //寫入Fail以及附加內容
                                excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Value = ScriptWorksheet.Cells[row, 6].Text + ", " + ScriptWorksheet.Cells[row, 7].Text+"\n";
                                WriteInLogString += " Sheet: " + excelWorksheet.Name + " Cell[" + excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].ToString().ToUpper() + "]:\n             " + ScriptWorksheet.Cells[row, 6].Text + ", " + ScriptWorksheet.Cells[row, 7].Text + "\n";
                                //script註解欄有值再寫入註解
                                if (ScriptWorksheet.Cells[row, 8].Value != null)
                                {
                                    if (excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment == null)
                                    {
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].AddComment(ScriptWorksheet.Cells[row, 8].Text+"\n", "User");
                                        WriteInLogString += "             新增註解: " + ScriptWorksheet.Cells[row, 8].Text + "\n";
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment.AutoFit = true;
                                    }
                                    else
                                    {
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment.Text = ScriptWorksheet.Cells[row, 8].Text+"\n";
                                        WriteInLogString += "             新增註解: " + ScriptWorksheet.Cells[row, 8].Text + "\n";
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment.AutoFit = true;
                                    }
                                }
                                //Note
                                if (ScriptWorksheet.Cells[row, 9].Value != null)
                                {
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].Value = ScriptWorksheet.Cells[row, 10].Text+"\n";
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].AutoFitColumns();
                                    WriteInLogString += "             Note[" + excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].ToString() + "]: " + ScriptWorksheet.Cells[row, 10].Text + "\n";
                                }
                                UsedFailString.Add(ScriptWorksheet.Cells[row, 6].Text);
                                excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].AutoFitColumns();
                                richTextBox1.AppendText(WriteInLogString);
                                continue;
                            }
                            //判斷欄位第一個字眼是否出現在fail字眼, 有則續寫
                            else if (UsedFailString.Contains(excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Text.Split(',')[0]))
                            {
                                //寫入Fail以及附加內容
                                excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Value += ScriptWorksheet.Cells[row, 7].Text+"\n";
                                WriteInLogString += " Sheet: " + excelWorksheet.Name + " Cell[" + excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].ToString().ToUpper() + "]:\n             " + ScriptWorksheet.Cells[row, 6].Text + ", " + ScriptWorksheet.Cells[row, 7].Text + "\n";
                                //script註解欄有值再寫入註解
                                if (ScriptWorksheet.Cells[row, 8].Value != null)
                                {
                                    if (excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment == null)
                                    {
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].AddComment(ScriptWorksheet.Cells[row, 8].Text+"\n", "User");
                                        WriteInLogString += "             新增註解: " + ScriptWorksheet.Cells[row, 8].Text + "\n";
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment.AutoFit = true;
                                    }
                                    else
                                    {
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment.Text += ScriptWorksheet.Cells[row, 8].Text+"\n";
                                        WriteInLogString += "             續寫註解: " + ScriptWorksheet.Cells[row, 8].Text + "\n";
                                        excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].Comment.AutoFit = true;
                                    }
                                }
                                //Note
                                if (ScriptWorksheet.Cells[row, 9].Value != null)
                                {
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].Value += ScriptWorksheet.Cells[row, 10].Text+"\n";
                                    excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].AutoFitColumns();
                                    WriteInLogString += "             Note[" + excelWorksheet.Cells[ScriptWorksheet.Cells[row, 9].Text].ToString() + "]: " + ScriptWorksheet.Cells[row, 10].Text + "\n";
                                }
                                UsedFailString.Add(ScriptWorksheet.Cells[row, 6].Text);
                                excelWorksheet.Cells[ScriptWorksheet.Cells[row, 3].Text].AutoFitColumns();
                                richTextBox1.AppendText(WriteInLogString);
                            }
                        }
                        //沒讀到Success或Fail
                        else
                        {
                            richTextBox1.AppendText("[Not written]" + DateTime.Now.ToString(" yyyyMMdd-HH:mm:ss ") + new FileInfo(htmltextBox.Text).Name + " No." + ScriptWorksheet.Cells[row, 1].Text + "無法判斷Success或Fail");
                        }
                    }

                    richTextBox1.AppendText("------------------------------------------------------------------------------------------------------\nFinished\n"+ time.ToString("yyyyMMdd-HH-mm-ss")+".log Saved.");
                    richTextBox1.SaveFile(@"log\" + time.ToString("yyyyMMdd-HH-mm-ss") + ".log", RichTextBoxStreamType.PlainText);
                    excelPackage.Save();
                }
            }
        }

        /// <summary>
        /// 檢查Script、Excel、Html、Html column
        /// </summary>
        /// <returns></returns>
        public bool CheckTextboxNull_or_NoSuchFile()
        {
            if (!File.Exists(scripttextBox.Text))
            {
                MessageBox.Show("請選擇正確的Script\n" + scripttextBox.Text);
                return false;
            }
            if (!File.Exists(exceltextBox.Text))
            {
                MessageBox.Show("請選擇正確的Excel\n" + exceltextBox.Text);
                return false;
            }
            if (!File.Exists(htmltextBox.Text))
            {
                MessageBox.Show("請選擇正確的HTML\n" + htmltextBox.Text);
                return false;
            }
            if (htmlcolumntextBox.Text == "" || int.Parse(htmlcolumntextBox.Text) < 0)
            {
                MessageBox.Show("請輸入正確Html Column");
                return false;
            }
            return true;
        }

        /// <summary>
        /// 輸入HTML路徑，回傳Datatable
        /// </summary>
        /// <param name="html"></param>
        /// <returns>Html Table</returns>
        public DataTable GetHtmlTable(string htmlpath)
        {
            string html_string;
            using (var fs = new FileStream(htmlpath, FileMode.Open, FileAccess.Read))
            {
                using (var sr = new StreamReader(fs))
                {
                    html_string = sr.ReadToEnd();
                }
            }

            var datatable = new DataTable();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html_string);
            foreach (HtmlNode table in doc.DocumentNode.SelectNodes("//table"))
            {
                foreach (HtmlNode row in table.SelectNodes("tr"))
                {
                    var headerCells = row.SelectNodes("th");
                    // add column name
                    if (headerCells != null)
                    {
                        for (int i = 0; i < headerCells.Count; i++)
                        {
                            datatable.Columns.Add(Convert.ToChar('A' + i).ToString());
                        }
                    }
                    // add th
                    if (headerCells != null)
                    {
                        var dataRow = datatable.NewRow();
                        for (int i = 0; i < headerCells.Count; i++)
                        {
                            dataRow[i] = headerCells[i].InnerText;
                        }
                        datatable.Rows.Add(dataRow);
                    }
                    // add td
                    var dataCells = row.SelectNodes("td");
                    if (dataCells != null)
                    {
                        var dataRow = datatable.NewRow();
                        for (int i = 0; i < dataCells.Count; i++)
                        {
                            dataRow[i] = dataCells[i].InnerText;
                        }
                        datatable.Rows.Add(dataRow);
                    }
                }
            }
            return datatable;
        }

        /// <summary>
        /// 按下開啟script按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (File.Exists(scripttextBox.Text))
            {
                System.Diagnostics.Process.Start(scripttextBox.Text);
            }
        }

        /// <summary>
        /// 按下開啟Excel按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (File.Exists(exceltextBox.Text))
            {
                System.Diagnostics.Process.Start(exceltextBox.Text);
            }
        }

        /// <summary>
        /// 按下開啟Html按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (File.Exists(htmltextBox.Text))
            {
                System.Diagnostics.Process.Start(htmltextBox.Text);
            }
        }

        /// <summary>
        /// 打開log資料夾
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Directory.GetCurrentDirectory() + @"\log"))
            {
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\log");
            }
            Process.Start(@"log");
        }
    }
}
