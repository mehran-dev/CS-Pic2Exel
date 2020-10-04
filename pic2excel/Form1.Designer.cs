namespace pic2excel
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
            this.pic1 = new System.Windows.Forms.PictureBox();
            this.btnLoad = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.lbl1 = new System.Windows.Forms.Label();
            this.pBar1 = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.pic1)).BeginInit();
            this.SuspendLayout();
            // 
            // pic1
            // 
            this.pic1.BackColor = System.Drawing.Color.Teal;
            this.pic1.Image = global::pic2excel.Properties.Resources._1_12186_firewatch_hd_firewatch_desktop_background;
            this.pic1.Location = new System.Drawing.Point(29, 12);
            this.pic1.Name = "pic1";
            this.pic1.Size = new System.Drawing.Size(434, 308);
            this.pic1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pic1.TabIndex = 0;
            this.pic1.TabStop = false;
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(43, 386);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(75, 23);
            this.btnLoad.TabIndex = 1;
            this.btnLoad.Text = "load picture";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(124, 386);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(144, 23);
            this.btnExcel.TabIndex = 2;
            this.btnExcel.Text = "Generate Excel File ";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.button2_Click);
            // 
            // lbl1
            // 
            this.lbl1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.lbl1.Location = new System.Drawing.Point(35, 352);
            this.lbl1.Name = "lbl1";
            this.lbl1.Size = new System.Drawing.Size(428, 23);
            this.lbl1.TabIndex = 3;
            this.lbl1.Text = "0";
            // 
            // pBar1
            // 
            this.pBar1.Location = new System.Drawing.Point(37, 415);
            this.pBar1.Name = "pBar1";
            this.pBar1.Size = new System.Drawing.Size(426, 10);
            this.pBar1.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(488, 450);
            this.Controls.Add(this.pBar1);
            this.Controls.Add(this.lbl1);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.pic1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.pic1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pic1;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Label lbl1;
        private System.Windows.Forms.ProgressBar pBar1;
    }
}

