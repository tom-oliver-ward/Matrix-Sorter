using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Matrix_Sorter
{
    class Processing
    {


        ReadExcel readExcel = new ReadExcel();
        internal void excelLoop(Form1 formObject)
        {
            Form1 form1 = formObject;
            int length = System.Convert.ToInt32(System.Convert.ToDouble(form1.SpreadSheets2Convert.Items.Count.ToString()));
            Application.DoEvents();
            for (int i = 0; i < length; i++)
            {
                form1.FilenumTB.Text = "Processing file " + (i + 1) + " of " + length;
                Application.DoEvents();
                readExcel.readFile(i, form1);
            }
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show("Complete", "", buttons);
        }

        internal void testIfDoneExcel(int j, int k, Excel.Cell cell)
        {
            if (j > 0 && k == 0 && cell == null)
            {                
                readExcel.writeCell = 1;               
            }
        }

        internal void transferExcelData(int j, int k, Excel.Cell cell, Form1 formObject)
        {
            Form1 form1 = formObject;
            if (j > 0 && k != 0 && readExcel.writeCell == 1)
            {             
                bool fileExists=false;
                if (k == 1) { readExcel.name = cell.Text; }
                if (k == 2) { readExcel.location = cell.Text; }
                if (readExcel.name != null && readExcel.location != null)
                { 
                    fileExists = TestIfDoneText(readExcel.name, readExcel.location); 
                }

                if (cell != null && fileExists== true) 
                { 
                    Matrix_Writer.writeCell(cell.Text, form1);                    
                }
                else 
                { 
                    Matrix_Writer.writeCell(" ", form1); 
                }
            }           
        }

        private bool TestIfDoneText(string name, string location)
        {
            bool value = File.Exists(string.Concat(location, "\\", name, "Sorted.txt"));
            return value;
        }

        internal void writeSpectra(Form1 formObject)
        {
            Form1 form1 = formObject;

            if (readExcel.name != null && readExcel.location != null)
            {
                FindSpectra.readWriteSpectra(readExcel.name, readExcel.location, form1);
                newline(form1);                
                readExcel.name = null; readExcel.location = null;
            }
        }

        private void newline(Form1 formObject)
        {
            Form1 form1 = formObject;
            using (StreamWriter sw3 = File.AppendText(form1.file))
            {
                sw3.WriteLine(string.Concat(" "));
                sw3.Close();
            }
        }
    }
}
