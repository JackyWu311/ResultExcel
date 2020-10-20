using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using ResultExcel.Class;
using ResultExcel.Component;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

using Excel = Microsoft.Office.Interop.Excel;

using Point = System.Drawing.Point;

namespace ResultExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            _CMS = GetCMS();
        }

        /// <summary>
        /// 點執行後檢查+寫入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WriteExcel_Click(object sender, EventArgs e)
        {
            //沒有tab不執行
            if (tabControl1.TabCount < 1)
                return;

            //執行前檢查FilePath是否能開啟
            if (!File.Exists(Currentblockpage().ExcelFilePath()) || !File.Exists(Currentblockpage().HtmlFilePath()))
            {
                MessageBox.Show("選擇正確的 Html 和 Excel 再執行");
                return;
            }

            //讀取指定的Html column並傳回dataTable
            DataTable dataTable = new DataTable();
            using (var fs = new FileStream(Currentblockpage().HtmlFilePath(), FileMode.Open, FileAccess.Read))
            {
                using (var sr = new StreamReader(fs))
                {
                    string html_string = sr.ReadToEnd();
                    dataTable = GetHtmlTable(html_string);
                }
            }

            //檢查htmlcolumn是否大於實際Column
            if (Currentblockpage().GetHtmlColumn() == "")
            {
                MessageBox.Show("未輸入HTML欄位");
                return;
            }
            else
            {
                if (int.Parse(Currentblockpage().GetHtmlColumn()) > dataTable.Columns.Count)
                {
                    MessageBox.Show("HTML欄位大於檔案欄位" + dataTable.Columns.Count);
                    return;
                }
            }

            //block數量不能為0
            List<Block> blocks = Currentblockpage().GetBlockList();
            if (blocks.Count < 1)
            {
                MessageBox.Show("請新增步驟");
                return;
            }

            //讀取Excel所有工作表並開始執行
            List<string> worksheetList = new List<string>();
            Excel.Application application = new Excel.Application
            {
                DisplayAlerts = false
            };
            Workbook workbook = application.Workbooks.Open(Currentblockpage().ExcelFilePath());
            try
            {
                //add list Excel所有工作表
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    worksheetList.Add(worksheet.Name);
                }

                int count = 1;
                //檢查各block
                foreach (Block block in blocks)
                {
                    //檢查block必填欄位
                    if (block.HtmlNo == "" || block.Sheet == "" || block.Success == "" || block.Fail == "" || block.Cell == "" || block.Note == "")
                    {
                        MessageBox.Show("第 " + count + " 區有空值未填入");
                        closeExcel(application, workbook);
                        return;
                    }

                    //檢查block 的htmlNo是否大於實際html Row
                    if (int.Parse(block.HtmlNo) + 1 > dataTable.Rows.Count)
                    {
                        MessageBox.Show("第 " + count + " 區HtmlNo超過檔案欄位");
                        closeExcel(application, workbook);
                        return;
                    }

                    //檢查block 的sheet是否存在Excel
                    if (!worksheetList.Contains(block.Sheet))
                    {
                        MessageBox.Show("第 " + count + " 區Sheet不在Excel中");
                        closeExcel(application, workbook);
                        return;
                    }
                    count++;
                }

                //檢查都通過則開始讀取寫入
                int column = int.Parse(Currentblockpage().GetHtmlColumn());
                List<string> HTMLSuccessString = new List<string> { "Success", "sucess", "Pass", "pass" };
                List<string> HTMLFailString = new List<string> { "Fail", "fail", "Skip", "skip", "Stop", "stop" };
                List<string> UsedSuccessString = new List<string>();  //紀錄寫入過的Success字眼
                List<string> UsedFailString = new List<string>(); //紀錄寫入過的Fail字眼
                foreach (Block block in blocks)
                {
                    //if html讀到success
                    if (HTMLSuccessString.Contains((string)dataTable.Rows[int.Parse(block.HtmlNo)][column - 1]))
                    {
                        Worksheet worksheet = workbook.Worksheets[block.Sheet]; //開啟對應worksheet
                        //如果是空欄位才寫入Success所有內容，否則跳過
                        if (worksheet.Range[block.Cell].Value2 == null)
                        {
                            worksheet.Range[block.Cell].Value2 = (block.Success + ", " + block.SuccessContent);
                            worksheet.Range[block.Cell].ClearComments();  //一定要清除註解，不然寫入會出錯
                            worksheet.Range[block.Cell].AddComment(block.Comment);
                            worksheet.Range[block.Cell].Comment.Shape.TextFrame.AutoSize = true;
                            worksheet.Range[block.Note].Value2 = (block.NoteContent);
                            UsedSuccessString.Add(block.Success); //將使用過的Success字眼存著
                        }
                    }
                    //if 讀到fail、skip、stop
                    else if (HTMLFailString.Contains((string)dataTable.Rows[int.Parse(block.HtmlNo)][column - 1]))
                    {
                        Worksheet worksheet = workbook.Worksheets[block.Sheet]; //開啟對應worksheet
                        //空欄位直接寫入 或 判斷欄位第一個字眼是否出現在success字眼，有則覆蓋
                        if (worksheet.Range[block.Cell].Value2 == null || UsedSuccessString.Contains(((string)worksheet.Range[block.Cell].Value2).Split(',')[0]))
                        {
                            worksheet.Range[block.Cell].Value2 = (block.Fail + ", " + block.FailContent);
                            worksheet.Range[block.Cell].ClearComments();  //一定要清除註解，不然寫入會出錯
                            worksheet.Range[block.Cell].AddComment(block.Comment);
                            worksheet.Range[block.Cell].Comment.Shape.TextFrame.AutoSize = true;
                            worksheet.Range[block.Note].Value2 = block.NoteContent;
                            UsedFailString.Add(block.Fail);
                        }
                        //判斷欄位第一個字眼是否出現在fail字眼, 有則續寫
                        else if (UsedFailString.Contains(((string)worksheet.Range[block.Cell].Value2).Split(',')[0]))
                        {
                            worksheet.Range[block.Cell].Value2 = (worksheet.Range[block.Cell].Value2 + "\n" + block.FailContent);//續寫
                            string temp = worksheet.Range[block.Cell].Comment.Text(); //先取出註解
                            worksheet.Range[block.Cell].ClearComments();  //清除註解，不然寫入會出錯
                            worksheet.Range[block.Cell].AddComment(temp + "\n" + block.Comment);//續寫
                            worksheet.Range[block.Cell].Comment.Shape.TextFrame.AutoSize = true;
                            worksheet.Range[block.Note].Value2 = worksheet.Range[block.Note].Value2 + "\n" + block.NoteContent;
                            UsedFailString.Add(block.Fail);
                        }
                    }
                }
                workbook.Save();
            }//try finish
            catch (Exception E)
            {
                closeExcel(application, workbook);
                MessageBox.Show(E.ToString() + "\nFail in Excel");
                return;
            }
            closeExcel(application, workbook);
            MessageBox.Show("Write to excel finished !");
            Process.Start(Currentblockpage().ExcelFilePath()); //執行完打開excel
        }

        /// <summary>
        /// Excel存檔並關閉Excel application
        /// </summary>
        /// <param name="application"></param>
        /// <param name="workbook"></param>
        public void closeExcel(Excel.Application application, Excel.Workbook workbook)
        {
            workbook.Save();
            workbook.Close();
            application.Quit();
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(application);
            Process[] excelProcesses = Process.GetProcessesByName("excel");
            foreach (Process p in excelProcesses)
            {
                // use MainWindowTitle to distinguish this excel process with other excel processes
                if (string.IsNullOrEmpty(p.MainWindowTitle))
                {
                    p.Kill();
                }
            }
            GC.Collect();
        }

        /// <summary>
        /// 傳入html內容，傳回datatable
        /// </summary>
        /// <param name="html"></param>
        /// <returns>Html Table</returns>
        public DataTable GetHtmlTable(string html)
        {
            var dt = new DataTable();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
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
                            dt.Columns.Add(Convert.ToChar('A' + i).ToString());
                        }
                    }
                    // add th
                    if (headerCells != null)
                    {
                        var dataRow = dt.NewRow();
                        for (int i = 0; i < headerCells.Count; i++)
                        {
                            dataRow[i] = headerCells[i].InnerText;
                        }
                        dt.Rows.Add(dataRow);
                    }
                    // add td
                    var dataCells = row.SelectNodes("td");
                    if (dataCells != null)
                    {
                        var dataRow = dt.NewRow();
                        for (int i = 0; i < dataCells.Count; i++)
                        {
                            dataRow[i] = dataCells[i].InnerText;
                        }
                        dt.Rows.Add(dataRow);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 選擇script讀檔並新增一頁tabpage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 開啟OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Select Script File"
            };
            string path = Directory.GetCurrentDirectory() + @"\Script";
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            dialog.InitialDirectory = Directory.GetCurrentDirectory() + @"\Script";
            dialog.Filter = "script|*.script";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                using (StreamReader r = new StreamReader(dialog.FileName))  //新增page
                {
                    TabPage tabPage = new TabPage();
                    blockpage blockpage = new blockpage
                    {
                        Dock = DockStyle.Fill
                    };
                    tabPage.Controls.Add(blockpage);
                    tabPage.Text = dialog.SafeFileName;
                    tabControl1.TabPages.Add(tabPage);
                    string json = r.ReadToEnd();
                    ExcelObject excelObject = JsonConvert.DeserializeObject<ExcelObject>(json);
                    blockpage.Load_JsonFile(excelObject.HtmlColumn, excelObject.ExcelFile, excelObject.HtmlFile, excelObject.blocks);
                    tabControl1.SelectedTab = tabPage;
                }
            }
        }

        /// <summary>
        /// 儲存Script
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 儲存SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "Save script file",
                Filter = "Script files (*.script)|*.script",
                RestoreDirectory = true
            };
            string path = Directory.GetCurrentDirectory() + @"\Script";
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            sfd.InitialDirectory = path;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                List<Block> blocks = Currentblockpage().GetBlockList();
                ExcelObject excelObject = new ExcelObject(Currentblockpage().GetHtmlColumn(), blocks, Currentblockpage().ExcelFilePath(), Currentblockpage().HtmlFilePath());
                using (StreamWriter file = File.CreateText(sfd.FileName))
                {
                    JsonSerializer serializer = new JsonSerializer
                    {
                        Formatting = Formatting.Indented
                    };
                    serializer.Serialize(file, excelObject);
                    tabControl1.SelectedTab.Text = Path.GetFileName(sfd.FileName);
                }
            }
        }

        /// <summary>
        /// 回傳SelectedTab的blockpage
        /// </summary>
        /// <returns></returns>
        private blockpage Currentblockpage()
        {
            return tabControl1.SelectedTab.Controls.OfType<blockpage>().Single();
        }

        /// <summary>
        /// 新增blockpage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 新增NToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TabPage tabPage = new TabPage();
            tabPage.Text = "NewScript";
            blockpage blockpage = new blockpage
            {
                Dock = DockStyle.Fill
            };
            tabPage.Controls.Add(blockpage);
            tabControl1.TabPages.Add(tabPage);
            tabControl1.SelectedTab = tabPage;
        }

        #region 右鍵關閉tab (請勿刪除)

        private ContextMenuStrip _CMS;
        private Point _lastClickPos;

        //右鍵關閉tab
        private void tabControl1_MouseClick(object sender, MouseEventArgs e)
        {
            base.OnMouseClick(e);
            if (e.Button == MouseButtons.Right)
            {
                _lastClickPos = Cursor.Position;
                _CMS.Show(Cursor.Position);
            }
        }

        //右鍵關閉tab
        private ContextMenuStrip GetCMS()
        {
            ContextMenuStrip CMS = new ContextMenuStrip();
            CMS.Items.Add("Close", null, new EventHandler(Item_Clicked));
            return CMS;
        }

        //右鍵關閉tab
        private void Item_Clicked(object sender, EventArgs e)
        {
            for (int i = 0; i < tabControl1.TabCount; i++)
            {
                System.Drawing.Rectangle rect = tabControl1.GetTabRect(i);
                if (rect.Contains(tabControl1.PointToClient(_lastClickPos)))
                {
                    tabControl1.TabPages.RemoveAt(i);
                }
            }
        }

        #endregion 右鍵關閉tab (請勿刪除)
    }
}