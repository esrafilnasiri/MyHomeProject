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
            this.btnOnlineCheckMaxZarar = new System.Windows.Forms.Button();
            this.btnOneMaket = new System.Windows.Forms.Button();
            this.txtMarketId = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnOnlineCheck
            // 
            this.btnOnlineCheck.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOnlineCheck.Location = new System.Drawing.Point(12, 12);
            this.btnOnlineCheck.Name = "btnOnlineCheck";
            this.btnOnlineCheck.Size = new System.Drawing.Size(75, 23);
            this.btnOnlineCheck.TabIndex = 0;
            this.btnOnlineCheck.Text = "button1";
            this.btnOnlineCheck.UseVisualStyleBackColor = true;
            this.btnOnlineCheck.Click += new System.EventHandler(this.BtnOnlineCheck_Click);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(12, 41);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(697, 320);
            this.textBox1.TabIndex = 1;
            // 
            // btnOnlineCheckMaxZarar
            // 
            this.btnOnlineCheckMaxZarar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOnlineCheckMaxZarar.Location = new System.Drawing.Point(93, 12);
            this.btnOnlineCheckMaxZarar.Name = "btnOnlineCheckMaxZarar";
            this.btnOnlineCheckMaxZarar.Size = new System.Drawing.Size(103, 23);
            this.btnOnlineCheckMaxZarar.TabIndex = 2;
            this.btnOnlineCheckMaxZarar.Text = "بررسی لحظه ای";
            this.btnOnlineCheckMaxZarar.UseVisualStyleBackColor = true;
            this.btnOnlineCheckMaxZarar.Click += new System.EventHandler(this.btnOnlineCheckMaxZarar_Click);
            // 
            // btnOneMaket
            // 
            this.btnOneMaket.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOneMaket.Location = new System.Drawing.Point(202, 12);
            this.btnOneMaket.Name = "btnOneMaket";
            this.btnOneMaket.Size = new System.Drawing.Size(75, 23);
            this.btnOneMaket.TabIndex = 3;
            this.btnOneMaket.Text = "یک مارکت";
            this.btnOneMaket.UseVisualStyleBackColor = true;
            this.btnOneMaket.Click += new System.EventHandler(this.btnOneMaket_Click);
            // 
            // txtMarketId
            // 
            this.txtMarketId.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMarketId.Location = new System.Drawing.Point(283, 13);
            this.txtMarketId.Name = "txtMarketId";
            this.txtMarketId.Size = new System.Drawing.Size(227, 21);
            this.txtMarketId.TabIndex = 4;
            this.txtMarketId.Text = "29974853866926823";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(721, 373);
            this.Controls.Add(this.txtMarketId);
            this.Controls.Add(this.btnOneMaket);
            this.Controls.Add(this.btnOnlineCheckMaxZarar);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnOnlineCheck);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOnlineCheck;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnOnlineCheckMaxZarar;
        private System.Windows.Forms.Button btnOneMaket;
        private System.Windows.Forms.TextBox txtMarketId;
    }
}

