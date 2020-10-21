using ResultExcel.Class;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ResultExcel.Component
{
    public partial class blockpage : UserControl
    {
        public blockpage()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 點Add按鈕新增第一個UserControlBlock
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddControlbutton_Click(object sender, EventArgs e)
        {
            UserControlBlock userControlBlock = new UserControlBlock();
            userControlBlock.ReturnBlock += UserControlBlock_ReturnBlock;
            userControlBlock.Enabled = true;
            flowLayoutPanel1.Controls.Add(userControlBlock);

            //Scroll to button
            Button Tempbutton = new Button();
            flowLayoutPanel1.Controls.Add(Tempbutton);
            flowLayoutPanel1.ScrollControlIntoView(Tempbutton);
            flowLayoutPanel1.Controls.Remove(Tempbutton);
            Tempbutton.Dispose();
        }

        /// <summary>
        /// 點複製按鈕
        /// </summary>
        /// <param name="obj"></param>
        private void UserControlBlock_ReturnBlock(Block obj)
        {
            UserControlBlock userControlBlock = new UserControlBlock();
            userControlBlock.SetBlock(obj);
            userControlBlock.ReturnBlock += UserControlBlock_ReturnBlock;
            flowLayoutPanel1.Controls.Add(userControlBlock);

            //Scroll to button
            Button Tempbutton = new Button();
            flowLayoutPanel1.Controls.Add(Tempbutton);
            flowLayoutPanel1.ScrollControlIntoView(Tempbutton);
            flowLayoutPanel1.Controls.Remove(Tempbutton);
            Tempbutton.Dispose();
        }

        /// <summary>
        /// 選擇HTML
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Htmlbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Html File";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Filter = "HTML|*.html";
            if (dialog.ShowDialog() == DialogResult.OK)
                HtmltextBox.Text = dialog.FileName;
        }

        /// <summary>
        /// 選擇Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Excelbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Excel File";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (dialog.ShowDialog() == DialogResult.OK)
                ExceltextBox.Text = dialog.FileName;
        }

        /// <summary>
        /// 更新layout
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void flowLayoutPanel1_ControlAdded(object sender, ControlEventArgs e)
        {
            if (flowLayoutPanel1.Controls.OfType<UserControlBlock>().Count() < 1)
            {
                Addbutton.Visible = true;
                Addbutton.Enabled = true;
            }
            //修改block id
            int controlCount = 1;
            foreach (UserControlBlock controlBlock in flowLayoutPanel1.Controls.OfType<UserControlBlock>())
            {
                controlBlock.setlabel2(controlCount);
                controlCount++;
            }
        }

        /// <summary>
        /// 更新layout
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void flowLayoutPanel1_ControlRemoved(object sender, ControlEventArgs e)
        {
            if (flowLayoutPanel1.Controls.OfType<UserControlBlock>().Count() < 1)
            {
                Addbutton.Visible = true;
                Addbutton.Enabled = true;
            }
            //修改block id
            int controlCount = 1;
            foreach (UserControlBlock controlBlock in flowLayoutPanel1.Controls.OfType<UserControlBlock>())
            {
                controlBlock.setlabel2(controlCount);
                controlCount++;
            }
        }

        /// <summary>
        /// 回傳Html Path
        /// </summary>
        /// <returns></returns>
        public string HtmlFilePath()
        {
            return HtmltextBox.Text;
        }

        /// <summary>
        /// 回傳Excel Path
        /// </summary>
        /// <returns></returns>
        public string ExcelFilePath()
        {
            return ExceltextBox.Text;
        }

        /// <summary>
        /// 回傳blockpage中所有block
        /// </summary>
        /// <returns></returns>
        public List<Block> GetBlockList()
        {
            List<Block> blocks = new List<Block>();
            foreach (UserControlBlock userControlBlock in flowLayoutPanel1.Controls.OfType<UserControlBlock>())
            {
                blocks.Add(userControlBlock.Getblock());
            }
            return blocks;
        }

        /// <summary>
        /// 回傳html欄位
        /// </summary>
        /// <returns></returns>
        public string GetHtmlColumn()
        {
            return HtmlColumnTextbox.Text;
        }

        /// <summary>
        /// 將開啟的script寫入blockpage
        /// </summary>
        /// <param name="HtmlColumn"></param>
        /// <param name="ExcelFile"></param>
        /// <param name="HtmlFile"></param>
        /// <param name="blocks"></param>
        public void Load_JsonFile(string HtmlColumn, string ExcelFile, string HtmlFile, List<Block> blocks)
        {
            HtmltextBox.Text = HtmlFile;
            HtmlColumnTextbox.Text = HtmlColumn;
            ExceltextBox.Text = ExcelFile;
            foreach (Block block in blocks)
            {
                UserControlBlock userControlBlock = new UserControlBlock();
                userControlBlock.SetBlock(block);
                userControlBlock.ReturnBlock += UserControlBlock_ReturnBlock;
                flowLayoutPanel1.Controls.Add(userControlBlock);
            }
        }

        /// <summary>
        /// HtmlColumn限定輸入數字
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HtmlColumnTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }

        /// <summary>
        /// 複製checked block
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CopyButton_Click(object sender, EventArgs e)
        {
            if (CopyNtimesTextBox.Text == "" || CopyNtimesTextBox.Text == "1")
            {
                foreach (UserControlBlock userControlBlock in flowLayoutPanel1.Controls.OfType<UserControlBlock>().Where(ucb => ucb.Selected == true))
                {
                    UserControlBlock_ReturnBlock(userControlBlock.Getblock());
                    userControlBlock.SetCheckbox(false);
                }
            }
            else
            {
                int count = int.Parse(CopyNtimesTextBox.Text);
                List<UserControlBlock> userControlBlocks = new List<UserControlBlock>();
                foreach (UserControlBlock userControlBlock in flowLayoutPanel1.Controls.OfType<UserControlBlock>().Where(ucb => ucb.Selected == true))
                {
                    userControlBlocks.Add(userControlBlock);
                    userControlBlock.SetCheckbox(false);
                }
                for (int i = 0; i < count; i++)
                {
                    foreach (UserControlBlock userControlBlock in userControlBlocks)
                    {
                        //只複製4個必填欄位
                        UserControlBlock_ReturnBlock(userControlBlock.GetBlock_Success_Fail_Sheet_Note());
                    }
                }
            }
        }

        /// <summary>
        /// 限定輸入數字
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CopyNtimesTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }

        /// <summary>
        /// 點2下打開HTML
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HtmltextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (File.Exists(HtmltextBox.Text))
            {
                System.Diagnostics.Process.Start(HtmltextBox.Text);
            }
            else
            {
                MessageBox.Show("找不到HTML\n" + HtmltextBox.Text);
            }
        }

        /// <summary>
        /// 點2下打開Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExceltextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (File.Exists(ExceltextBox.Text))
            {
                System.Diagnostics.Process.Start(ExceltextBox.Text);
            }
            else
            {
                MessageBox.Show("找不到Excel\n" + ExceltextBox.Text);
            }
        }
    }
}