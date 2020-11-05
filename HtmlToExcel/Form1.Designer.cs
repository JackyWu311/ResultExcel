namespace HtmlToExcel
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.htmltextBox = new System.Windows.Forms.TextBox();
            this.exceltextBox = new System.Windows.Forms.TextBox();
            this.scripttextBox = new System.Windows.Forms.TextBox();
            this.htmlcolumntextBox = new System.Windows.Forms.TextBox();
            this.scriptbutton = new System.Windows.Forms.Button();
            this.excelbutton = new System.Windows.Forms.Button();
            this.hmtlbutton = new System.Windows.Forms.Button();
            this.runbutton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // htmltextBox
            // 
            this.htmltextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.htmltextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.htmltextBox.Location = new System.Drawing.Point(91, 133);
            this.htmltextBox.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.htmltextBox.Name = "htmltextBox";
            this.htmltextBox.ReadOnly = true;
            this.htmltextBox.Size = new System.Drawing.Size(495, 27);
            this.htmltextBox.TabIndex = 0;
            // 
            // exceltextBox
            // 
            this.exceltextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.exceltextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.exceltextBox.Location = new System.Drawing.Point(92, 84);
            this.exceltextBox.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.exceltextBox.Name = "exceltextBox";
            this.exceltextBox.ReadOnly = true;
            this.exceltextBox.Size = new System.Drawing.Size(494, 27);
            this.exceltextBox.TabIndex = 1;
            // 
            // scripttextBox
            // 
            this.scripttextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.scripttextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.scripttextBox.Location = new System.Drawing.Point(92, 11);
            this.scripttextBox.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.scripttextBox.Name = "scripttextBox";
            this.scripttextBox.ReadOnly = true;
            this.scripttextBox.Size = new System.Drawing.Size(494, 27);
            this.scripttextBox.TabIndex = 2;
            // 
            // htmlcolumntextBox
            // 
            this.htmlcolumntextBox.Location = new System.Drawing.Point(111, 174);
            this.htmlcolumntextBox.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.htmlcolumntextBox.Name = "htmlcolumntextBox";
            this.htmlcolumntextBox.Size = new System.Drawing.Size(50, 27);
            this.htmlcolumntextBox.TabIndex = 3;
            this.htmlcolumntextBox.Text = "7";
            this.htmlcolumntextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.htmlcolumntextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox4_KeyPress);
            // 
            // scriptbutton
            // 
            this.scriptbutton.Location = new System.Drawing.Point(6, 7);
            this.scriptbutton.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.scriptbutton.Name = "scriptbutton";
            this.scriptbutton.Size = new System.Drawing.Size(76, 37);
            this.scriptbutton.TabIndex = 4;
            this.scriptbutton.Text = "Script";
            this.scriptbutton.UseVisualStyleBackColor = true;
            this.scriptbutton.Click += new System.EventHandler(this.scriptbutton_Click);
            // 
            // excelbutton
            // 
            this.excelbutton.Location = new System.Drawing.Point(6, 77);
            this.excelbutton.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.excelbutton.Name = "excelbutton";
            this.excelbutton.Size = new System.Drawing.Size(76, 37);
            this.excelbutton.TabIndex = 5;
            this.excelbutton.Text = "Excel";
            this.excelbutton.UseVisualStyleBackColor = true;
            this.excelbutton.Click += new System.EventHandler(this.excelbutton_Click);
            // 
            // hmtlbutton
            // 
            this.hmtlbutton.Location = new System.Drawing.Point(5, 126);
            this.hmtlbutton.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.hmtlbutton.Name = "hmtlbutton";
            this.hmtlbutton.Size = new System.Drawing.Size(76, 37);
            this.hmtlbutton.TabIndex = 6;
            this.hmtlbutton.Text = "HTML";
            this.hmtlbutton.UseVisualStyleBackColor = true;
            this.hmtlbutton.Click += new System.EventHandler(this.hmtlbutton_Click);
            // 
            // runbutton
            // 
            this.runbutton.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.runbutton.Location = new System.Drawing.Point(555, 167);
            this.runbutton.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.runbutton.Name = "runbutton";
            this.runbutton.Size = new System.Drawing.Size(94, 37);
            this.runbutton.TabIndex = 7;
            this.runbutton.Text = "Run";
            this.runbutton.UseVisualStyleBackColor = true;
            this.runbutton.Click += new System.EventHandler(this.runbutton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 177);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 16);
            this.label1.TabIndex = 8;
            this.label1.Text = "Html Column";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.scriptbutton);
            this.panel1.Controls.Add(this.runbutton);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.htmlcolumntextBox);
            this.panel1.Controls.Add(this.scripttextBox);
            this.panel1.Controls.Add(this.excelbutton);
            this.panel1.Controls.Add(this.hmtlbutton);
            this.panel1.Controls.Add(this.htmltextBox);
            this.panel1.Controls.Add(this.exceltextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(654, 208);
            this.panel1.TabIndex = 9;
            // 
            // button4
            // 
            this.button4.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button4.Location = new System.Drawing.Point(452, 167);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(95, 37);
            this.button4.TabIndex = 14;
            this.button4.Text = "Log Folder";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(594, 128);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(55, 32);
            this.button3.TabIndex = 13;
            this.button3.Text = "Open";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.Location = new System.Drawing.Point(594, 79);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(55, 32);
            this.button2.TabIndex = 12;
            this.button2.Text = "Open";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(594, 9);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(55, 32);
            this.button1.TabIndex = 11;
            this.button1.Text = "Open";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(35, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 16);
            this.label2.TabIndex = 10;
            this.label2.Text = "Sheet";
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.ImeMode = System.Windows.Forms.ImeMode.Close;
            this.comboBox1.Location = new System.Drawing.Point(92, 50);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(494, 24);
            this.comboBox1.TabIndex = 9;
            // 
            // richTextBox1
            // 
            this.richTextBox1.AutoWordSelection = true;
            this.richTextBox1.BackColor = System.Drawing.SystemColors.Control;
            this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.richTextBox1.Location = new System.Drawing.Point(0, 208);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(654, 448);
            this.richTextBox1.TabIndex = 10;
            this.richTextBox1.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(654, 656);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.Name = "Form1";
            this.Text = "Result 回填";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox htmltextBox;
        private System.Windows.Forms.TextBox exceltextBox;
        private System.Windows.Forms.TextBox scripttextBox;
        private System.Windows.Forms.TextBox htmlcolumntextBox;
        private System.Windows.Forms.Button scriptbutton;
        private System.Windows.Forms.Button excelbutton;
        private System.Windows.Forms.Button hmtlbutton;
        private System.Windows.Forms.Button runbutton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button4;
    }
}

