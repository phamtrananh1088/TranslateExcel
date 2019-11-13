namespace Anh.Tuhoconline
{
    partial class Form1
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
            this.btn共通単語 = new System.Windows.Forms.Button();
            this.btn1000共通単語 = new System.Windows.Forms.Button();
            this.btn2000共通単語 = new System.Windows.Forms.Button();
            this.btnNguphapN5 = new System.Windows.Forms.Button();
            this.btnNguphapN4 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn共通単語
            // 
            this.btn共通単語.Location = new System.Drawing.Point(12, 12);
            this.btn共通単語.Name = "btn共通単語";
            this.btn共通単語.Size = new System.Drawing.Size(98, 23);
            this.btn共通単語.TabIndex = 0;
            this.btn共通単語.Text = "3000共通単語";
            this.btn共通単語.UseVisualStyleBackColor = true;
            this.btn共通単語.Click += new System.EventHandler(this.btn共通単語_Click);
            // 
            // btn1000共通単語
            // 
            this.btn1000共通単語.Location = new System.Drawing.Point(12, 41);
            this.btn1000共通単語.Name = "btn1000共通単語";
            this.btn1000共通単語.Size = new System.Drawing.Size(98, 23);
            this.btn1000共通単語.TabIndex = 1;
            this.btn1000共通単語.Text = "1000共通単語";
            this.btn1000共通単語.UseVisualStyleBackColor = true;
            this.btn1000共通単語.Click += new System.EventHandler(this.btn1000共通単語_Click);
            // 
            // btn2000共通単語
            // 
            this.btn2000共通単語.Location = new System.Drawing.Point(12, 70);
            this.btn2000共通単語.Name = "btn2000共通単語";
            this.btn2000共通単語.Size = new System.Drawing.Size(98, 23);
            this.btn2000共通単語.TabIndex = 2;
            this.btn2000共通単語.Text = "2000共通単語";
            this.btn2000共通単語.UseVisualStyleBackColor = true;
            this.btn2000共通単語.Click += new System.EventHandler(this.btn2000共通単語_Click);
            // 
            // btnNguphapN5
            // 
            this.btnNguphapN5.Location = new System.Drawing.Point(13, 100);
            this.btnNguphapN5.Name = "btnNguphapN5";
            this.btnNguphapN5.Size = new System.Drawing.Size(97, 23);
            this.btnNguphapN5.TabIndex = 3;
            this.btnNguphapN5.Text = "Ngu phap N5";
            this.btnNguphapN5.UseVisualStyleBackColor = true;
            this.btnNguphapN5.Click += new System.EventHandler(this.btnNguphapN5_Click);
            // 
            // btnNguphapN4
            // 
            this.btnNguphapN4.Location = new System.Drawing.Point(13, 129);
            this.btnNguphapN4.Name = "btnNguphapN4";
            this.btnNguphapN4.Size = new System.Drawing.Size(97, 23);
            this.btnNguphapN4.TabIndex = 4;
            this.btnNguphapN4.Text = "Ngu phap N4";
            this.btnNguphapN4.UseVisualStyleBackColor = true;
            this.btnNguphapN4.Click += new System.EventHandler(this.btnNguphapN4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.btnNguphapN4);
            this.Controls.Add(this.btnNguphapN5);
            this.Controls.Add(this.btn2000共通単語);
            this.Controls.Add(this.btn1000共通単語);
            this.Controls.Add(this.btn共通単語);
            this.Name = "Form1";
            this.Text = "Tuhoconline";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn共通単語;
        private System.Windows.Forms.Button btn1000共通単語;
        private System.Windows.Forms.Button btn2000共通単語;
        private System.Windows.Forms.Button btnNguphapN5;
        private System.Windows.Forms.Button btnNguphapN4;
    }
}

