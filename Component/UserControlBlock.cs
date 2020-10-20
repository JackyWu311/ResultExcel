using ResultExcel.Class;
using System;
using System.Windows.Forms;

namespace ResultExcel.Component
{
    public partial class UserControlBlock : UserControl
    {
        public event Action<Block> ReturnBlock;

        public UserControlBlock()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 回傳 UserControlBlock 欄位的值
        /// </summary>
        /// <returns>all block content</returns>
        public Block Getblock()
        {
            Block block = new Block
            {
                HtmlNo = textBox1.Text,
                Success = textBox2.Text,
                SuccessContent = richTextBox1.Text,
                Fail = textBox3.Text,
                FailContent = richTextBox2.Text,
                Sheet = textBox4.Text,
                Cell = textBox5.Text,
                Comment = richTextBox3.Text,
                Note = textBox6.Text,
                NoteContent = richTextBox4.Text,
            };
            return block;
        }

        /// <summary>
        /// 複製或開啟檔案時寫入block
        /// </summary>
        /// <param name="block"></param>
        public void SetBlock(Block block)
        {
            textBox1.Text = block.HtmlNo;
            textBox2.Text = block.Success;
            richTextBox1.Text = block.SuccessContent;
            textBox3.Text = block.Fail;
            richTextBox2.Text = block.FailContent;
            textBox4.Text = block.Sheet;
            textBox5.Text = block.Cell;
            richTextBox3.Text = block.Comment;
            textBox6.Text = block.Note;
            richTextBox4.Text = block.NoteContent;
        }

        /// <summary>
        /// 點複製按鈕，複製四個項目
        /// </summary>
        /// <returns>Success、Fail、Sheet、Note</returns>
        public Block GetBlock_Success_Fail_Sheet_Note()
        {
            Block block = new Block
            {
                Success = textBox2.Text,
                Fail = textBox3.Text,
                Sheet = textBox4.Text,
                Note = textBox6.Text,
            };
            return block;
        }

        // 限定輸入數字
        private void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }

        // 複製user control
        private void Button1_Click(object sender, EventArgs e)
        {
            ReturnBlock.Invoke(GetBlock_Success_Fail_Sheet_Note());
        }

        //刪除 usercontrolblock
        private void button3_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        /// <summary>
        /// 改UserControlBlock編號
        /// </summary>
        /// <param name="id"></param>
        public void setlabel2(int id)
        {
            label2.Text = id.ToString();
        }

        public bool Selected = false;

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Selected = checkBox1.Checked;
        }

        /// <summary>
        /// 設定CheckBox是否勾選
        /// </summary>
        /// <param name="Selected"></param>
        public void SetCheckbox(bool Selected)
        {
            checkBox1.Checked = Selected;
        }
    }
}