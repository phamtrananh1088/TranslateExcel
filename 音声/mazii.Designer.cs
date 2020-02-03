namespace Anh.音声
{
    partial class mazii
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.cbSheetName = new System.Windows.Forms.ComboBox();
            this.btntokenizer = new System.Windows.Forms.Button();
            this.linkFileName = new System.Windows.Forms.LinkLabel();
            this.txtExcelName = new System.Windows.Forms.TextBox();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnkanji = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(2, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Sheet name：";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // cbSheetName
            // 
            this.cbSheetName.FormattingEnabled = true;
            this.cbSheetName.Location = new System.Drawing.Point(80, 48);
            this.cbSheetName.Name = "cbSheetName";
            this.cbSheetName.Size = new System.Drawing.Size(121, 20);
            this.cbSheetName.TabIndex = 3;
            // 
            // btntokenizer
            // 
            this.btntokenizer.Location = new System.Drawing.Point(80, 74);
            this.btntokenizer.Name = "btntokenizer";
            this.btntokenizer.Size = new System.Drawing.Size(75, 23);
            this.btntokenizer.TabIndex = 4;
            this.btntokenizer.Text = "tokenizer";
            this.btntokenizer.UseVisualStyleBackColor = true;
            this.btntokenizer.Click += new System.EventHandler(this.btntokenizer_Click);
            // 
            // linkFileName
            // 
            this.linkFileName.AutoSize = true;
            this.linkFileName.Location = new System.Drawing.Point(12, 23);
            this.linkFileName.Name = "linkFileName";
            this.linkFileName.Size = new System.Drawing.Size(61, 12);
            this.linkFileName.TabIndex = 0;
            this.linkFileName.TabStop = true;
            this.linkFileName.Text = "File name：";
            this.linkFileName.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.linkFileName.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFileName_LinkClicked);
            // 
            // txtExcelName
            // 
            this.txtExcelName.Location = new System.Drawing.Point(80, 23);
            this.txtExcelName.Name = "txtExcelName";
            this.txtExcelName.ReadOnly = true;
            this.txtExcelName.Size = new System.Drawing.Size(359, 19);
            this.txtExcelName.TabIndex = 1;
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.AutoToolTip = true;
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            this.toolStripProgressBar1.Step = 1;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1,
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2});
            this.statusStrip1.Location = new System.Drawing.Point(0, 239);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(502, 22);
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(0, 17);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnkanji
            // 
            this.btnkanji.Location = new System.Drawing.Point(80, 103);
            this.btnkanji.Name = "btnkanji";
            this.btnkanji.Size = new System.Drawing.Size(75, 23);
            this.btnkanji.TabIndex = 5;
            this.btnkanji.Text = "kanji";
            this.btnkanji.UseVisualStyleBackColor = true;
            this.btnkanji.Click += new System.EventHandler(this.btnkanji_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(80, 132);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(121, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "search -> example";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnsearchExample_Click);
            // 
            // mazii
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(502, 261);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbSheetName);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnkanji);
            this.Controls.Add(this.btntokenizer);
            this.Controls.Add(this.linkFileName);
            this.Controls.Add(this.txtExcelName);
            this.Controls.Add(this.statusStrip1);
            this.Name = "mazii";
            this.Text = "mazii";
            this.Load += new System.EventHandler(this.mazii_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbSheetName;
        private System.Windows.Forms.Button btntokenizer;
        private System.Windows.Forms.LinkLabel linkFileName;
        private System.Windows.Forms.TextBox txtExcelName;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnkanji;
        private System.Windows.Forms.Button button1;
	}
}

