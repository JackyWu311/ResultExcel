namespace ResultExcel.Component
{
    partial class blockpage
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

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.CopyNtimesTextBox = new System.Windows.Forms.TextBox();
            this.CopyButton = new System.Windows.Forms.Button();
            this.Addbutton = new System.Windows.Forms.Button();
            this.HtmlColumnTextbox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.HtmltextBox = new System.Windows.Forms.TextBox();
            this.Excelbutton = new System.Windows.Forms.Button();
            this.ExceltextBox = new System.Windows.Forms.TextBox();
            this.Htmlbutton = new System.Windows.Forms.Button();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightSkyBlue;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.CopyNtimesTextBox);
            this.panel1.Controls.Add(this.CopyButton);
            this.panel1.Controls.Add(this.Addbutton);
            this.panel1.Controls.Add(this.HtmlColumnTextbox);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.HtmltextBox);
            this.panel1.Controls.Add(this.Excelbutton);
            this.panel1.Controls.Add(this.ExceltextBox);
            this.panel1.Controls.Add(this.Htmlbutton);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(803, 62);
            this.panel1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(697, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "複製N次";
            // 
            // CopyNtimesTextBox
            // 
            this.CopyNtimesTextBox.AllowDrop = true;
            this.CopyNtimesTextBox.Location = new System.Drawing.Point(697, 32);
            this.CopyNtimesTextBox.Name = "CopyNtimesTextBox";
            this.CopyNtimesTextBox.Size = new System.Drawing.Size(64, 22);
            this.CopyNtimesTextBox.TabIndex = 6;
            this.CopyNtimesTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CopyNtimesTextBox_KeyPress);
            // 
            // CopyButton
            // 
            this.CopyButton.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.CopyButton.Location = new System.Drawing.Point(598, 31);
            this.CopyButton.Name = "CopyButton";
            this.CopyButton.Size = new System.Drawing.Size(93, 25);
            this.CopyButton.TabIndex = 0;
            this.CopyButton.Text = "多選複製";
            this.CopyButton.UseVisualStyleBackColor = true;
            this.CopyButton.Click += new System.EventHandler(this.CopyButton_Click);
            // 
            // Addbutton
            // 
            this.Addbutton.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Addbutton.Location = new System.Drawing.Point(505, 31);
            this.Addbutton.Name = "Addbutton";
            this.Addbutton.Size = new System.Drawing.Size(87, 25);
            this.Addbutton.TabIndex = 0;
            this.Addbutton.Text = "新增步驟";
            this.Addbutton.UseVisualStyleBackColor = true;
            this.Addbutton.Click += new System.EventHandler(this.AddControlbutton_Click);
            // 
            // HtmlColumnTextbox
            // 
            this.HtmlColumnTextbox.AllowDrop = true;
            this.HtmlColumnTextbox.Location = new System.Drawing.Point(576, 3);
            this.HtmlColumnTextbox.Name = "HtmlColumnTextbox";
            this.HtmlColumnTextbox.Size = new System.Drawing.Size(100, 22);
            this.HtmlColumnTextbox.TabIndex = 5;
            this.HtmlColumnTextbox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.HtmlColumnTextbox_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(503, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "HTML欄位";
            // 
            // HtmltextBox
            // 
            this.HtmltextBox.AllowDrop = true;
            this.HtmltextBox.Location = new System.Drawing.Point(3, 3);
            this.HtmltextBox.Name = "HtmltextBox";
            this.HtmltextBox.ReadOnly = true;
            this.HtmltextBox.Size = new System.Drawing.Size(413, 22);
            this.HtmltextBox.TabIndex = 0;
            this.HtmltextBox.TabStop = false;
            this.HtmltextBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.HtmltextBox_MouseDoubleClick);
            // 
            // Excelbutton
            // 
            this.Excelbutton.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Excelbutton.Location = new System.Drawing.Point(422, 32);
            this.Excelbutton.Name = "Excelbutton";
            this.Excelbutton.Size = new System.Drawing.Size(75, 25);
            this.Excelbutton.TabIndex = 3;
            this.Excelbutton.TabStop = false;
            this.Excelbutton.Text = "Excel";
            this.Excelbutton.UseVisualStyleBackColor = true;
            this.Excelbutton.Click += new System.EventHandler(this.Excelbutton_Click);
            // 
            // ExceltextBox
            // 
            this.ExceltextBox.AllowDrop = true;
            this.ExceltextBox.Location = new System.Drawing.Point(3, 31);
            this.ExceltextBox.Name = "ExceltextBox";
            this.ExceltextBox.ReadOnly = true;
            this.ExceltextBox.Size = new System.Drawing.Size(413, 22);
            this.ExceltextBox.TabIndex = 1;
            this.ExceltextBox.TabStop = false;
            this.ExceltextBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.ExceltextBox_MouseDoubleClick);
            // 
            // Htmlbutton
            // 
            this.Htmlbutton.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Htmlbutton.Location = new System.Drawing.Point(422, 4);
            this.Htmlbutton.Name = "Htmlbutton";
            this.Htmlbutton.Size = new System.Drawing.Size(75, 25);
            this.Htmlbutton.TabIndex = 2;
            this.Htmlbutton.TabStop = false;
            this.Htmlbutton.Text = "HTML";
            this.Htmlbutton.UseVisualStyleBackColor = true;
            this.Htmlbutton.Click += new System.EventHandler(this.Htmlbutton_Click);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoScroll = true;
            this.flowLayoutPanel1.BackColor = System.Drawing.Color.White;
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 62);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(803, 513);
            this.flowLayoutPanel1.TabIndex = 5;
            this.flowLayoutPanel1.ControlAdded += new System.Windows.Forms.ControlEventHandler(this.flowLayoutPanel1_ControlAdded);
            this.flowLayoutPanel1.ControlRemoved += new System.Windows.Forms.ControlEventHandler(this.flowLayoutPanel1_ControlRemoved);
            // 
            // blockpage
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.panel1);
            this.Name = "blockpage";
            this.Size = new System.Drawing.Size(803, 575);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox HtmltextBox;
        private System.Windows.Forms.Button Excelbutton;
        private System.Windows.Forms.TextBox ExceltextBox;
        private System.Windows.Forms.Button Htmlbutton;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button Addbutton;
        private System.Windows.Forms.TextBox HtmlColumnTextbox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button CopyButton;
        private System.Windows.Forms.TextBox CopyNtimesTextBox;
        private System.Windows.Forms.Label label2;
    }
}
