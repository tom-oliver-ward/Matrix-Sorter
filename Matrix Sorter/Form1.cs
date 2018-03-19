using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Matrix_Sorter
{
    public partial class Form1 : Form
    {
        Processing processing = new Processing();
        public string file;
        public string fileError;
        public string name = null;
        public string location = null;
        public List<string> missingSpectra = new List<string>();

        public Form1()
        {
            InitializeComponent();
            this.SpreadSheets2Convert.DragDrop += new
            System.Windows.Forms.DragEventHandler(this.SpreadSheets2Convert_DragDrop);
            this.SpreadSheets2Convert.DragEnter += new
            System.Windows.Forms.DragEventHandler(this.SpreadSheets2Convert_DragEnter);
            buttonConvert.Enabled = false;
        }

        private void SpreadSheets2Convert_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            int i;
            for (i = 0; i < s.Length; i++)
                SpreadSheets2Convert.Items.Add(s[i]);
        }

        private void SpreadSheets2Convert_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            processing.FileErrorPath(this);
            File.Delete(fileError);
            processing.excelLoop(this);
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show("\t\tComplete \nAny list of missed Spectra will open in Notepad", "", buttons);
            textBoxExcelLine.Text = "Complete";
            FilenumTB.Text = "Complete";

            try
            {
                Process.Start(fileError);
            }
            catch (System.ComponentModel.Win32Exception)
            {

            }
        }

        private void buttonMatrixF_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
            }
            textBoxMatrixFile.Text = file;
            buttonConvert.Enabled = true;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button_Clear_Click(object sender, EventArgs e)
        {
            SpreadSheets2Convert.Items.Clear();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
    }
}
