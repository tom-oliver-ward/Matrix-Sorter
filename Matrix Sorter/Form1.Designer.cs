﻿namespace Matrix_Sorter
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.FilenumTB = new System.Windows.Forms.TextBox();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.SpreadSheets2Convert = new System.Windows.Forms.ListBox();
            this.buttonMatrixF = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(2, 6);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(123, 20);
            this.textBox1.TabIndex = 8;
            this.textBox1.Text = "Drag Spreadsheet Here";
            // 
            // FilenumTB
            // 
            this.FilenumTB.Location = new System.Drawing.Point(273, 204);
            this.FilenumTB.Name = "FilenumTB";
            this.FilenumTB.Size = new System.Drawing.Size(140, 20);
            this.FilenumTB.TabIndex = 7;
            // 
            // buttonConvert
            // 
            this.buttonConvert.Location = new System.Drawing.Point(273, 174);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Size = new System.Drawing.Size(140, 23);
            this.buttonConvert.TabIndex = 6;
            this.buttonConvert.Text = "Convert";
            this.buttonConvert.UseVisualStyleBackColor = true;
            this.buttonConvert.Click += new System.EventHandler(this.buttonConvert_Click);
            // 
            // SpreadSheets2Convert
            // 
            this.SpreadSheets2Convert.AllowDrop = true;
            this.SpreadSheets2Convert.FormattingEnabled = true;
            this.SpreadSheets2Convert.Location = new System.Drawing.Point(2, 32);
            this.SpreadSheets2Convert.Name = "SpreadSheets2Convert";
            this.SpreadSheets2Convert.Size = new System.Drawing.Size(700, 95);
            this.SpreadSheets2Convert.TabIndex = 5;
            // 
            // buttonMatrixF
            // 
            this.buttonMatrixF.Location = new System.Drawing.Point(273, 138);
            this.buttonMatrixF.Name = "buttonMatrixF";
            this.buttonMatrixF.Size = new System.Drawing.Size(140, 23);
            this.buttonMatrixF.TabIndex = 9;
            this.buttonMatrixF.Text = "Select Matrix File";
            this.buttonMatrixF.UseVisualStyleBackColor = true;
            this.buttonMatrixF.Click += new System.EventHandler(this.buttonMatrixF_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(709, 292);
            this.Controls.Add(this.buttonMatrixF);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.FilenumTB);
            this.Controls.Add(this.buttonConvert);
            this.Controls.Add(this.SpreadSheets2Convert);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button buttonConvert;
        public System.Windows.Forms.ListBox SpreadSheets2Convert;
        private System.Windows.Forms.Button buttonMatrixF;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        public System.Windows.Forms.TextBox FilenumTB;
    }
}

