namespace Anh.音声
{
    partial class 音声
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
            this.btn聞く = new System.Windows.Forms.Button();
            this.linkFileName = new System.Windows.Forms.LinkLabel();
            this.txtExcelName = new System.Windows.Forms.TextBox();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "Sheet name：";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // cbSheetName
            // 
            this.cbSheetName.FormattingEnabled = true;
            this.cbSheetName.Location = new System.Drawing.Point(87, 31);
            this.cbSheetName.Name = "cbSheetName";
            this.cbSheetName.Size = new System.Drawing.Size(121, 20);
            this.cbSheetName.TabIndex = 9;
            // 
            // btn聞く
            // 
            this.btn聞く.Location = new System.Drawing.Point(87, 57);
            this.btn聞く.Name = "btn聞く";
            this.btn聞く.Size = new System.Drawing.Size(75, 23);
            this.btn聞く.TabIndex = 10;
            this.btn聞く.Text = "音声を聞く";
            this.btn聞く.UseVisualStyleBackColor = true;
            this.btn聞く.Click += new System.EventHandler(this.btn聞く_Click);
            // 
            // linkFileName
            // 
            this.linkFileName.AutoSize = true;
            this.linkFileName.Location = new System.Drawing.Point(19, 6);
            this.linkFileName.Name = "linkFileName";
            this.linkFileName.Size = new System.Drawing.Size(61, 12);
            this.linkFileName.TabIndex = 6;
            this.linkFileName.TabStop = true;
            this.linkFileName.Text = "File name：";
            this.linkFileName.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.linkFileName.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFileName_LinkClicked);
            // 
            // txtExcelName
            // 
            this.txtExcelName.Location = new System.Drawing.Point(87, 6);
            this.txtExcelName.Name = "txtExcelName";
            this.txtExcelName.ReadOnly = true;
            this.txtExcelName.Size = new System.Drawing.Size(359, 19);
            this.txtExcelName.TabIndex = 7;
            // 
            // toolStripProgressBar1
            // 
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
            this.statusStrip1.TabIndex = 11;
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
            // 音声
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(502, 261);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbSheetName);
            this.Controls.Add(this.btn聞く);
            this.Controls.Add(this.linkFileName);
            this.Controls.Add(this.txtExcelName);
            this.Controls.Add(this.statusStrip1);
            this.Name = "音声";
            this.Text = "音声";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbSheetName;
        private System.Windows.Forms.Button btn聞く;
        private System.Windows.Forms.LinkLabel linkFileName;
        private System.Windows.Forms.TextBox txtExcelName;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
	}
}

