namespace OnlineCheck
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
            this.btnOnlineCheck = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnOnlineCheck
            // 
            this.btnOnlineCheck.Location = new System.Drawing.Point(12, 12);
            this.btnOnlineCheck.Name = "btnOnlineCheck";
            this.btnOnlineCheck.Size = new System.Drawing.Size(75, 23);
            this.btnOnlineCheck.TabIndex = 0;
            this.btnOnlineCheck.Text = "button1";
            this.btnOnlineCheck.UseVisualStyleBackColor = true;
            this.btnOnlineCheck.Click += new System.EventHandler(this.btnOnlineCheck_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 41);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(697, 320);
            this.textBox1.TabIndex = 1;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(721, 373);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnOnlineCheck);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOnlineCheck;
        private System.Windows.Forms.TextBox textBox1;
    }
}

