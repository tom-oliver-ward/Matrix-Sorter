using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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

        public Form1()
        {
            InitializeComponent();
            this.SpreadSheets2Convert.DragDrop += new
            System.Windows.Forms.DragEventHandler(this.SpreadSheets2Convert_DragDrop);
            this.SpreadSheets2Convert.DragEnter += new
            System.Windows.Forms.DragEventHandler(this.SpreadSheets2Convert_DragEnter);
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
            processing.excelLoop(this);            
        }

        private void buttonMatrixF_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
