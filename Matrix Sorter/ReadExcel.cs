using Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;


namespace Matrix_Sorter
{
    class ReadExcel
    {
        DataTransfer dataTransfer = new DataTransfer();
        Matrix_Writer matrixWriter = new Matrix_Writer();
        
             
        public int writeCell;
        

        public void readFile(int i, Form1 formObject)
        {
            Form1 form1 = formObject;
            string filename = (string)form1.SpreadSheets2Convert.Items[i];
            int j = 0;
            
            foreach (var worksheet in Workbook.Worksheets(filename))
            {
                foreach (var row in worksheet.Rows)                
                {
                    form1.textBoxExcelLine.Text = "Processing line " + (j + 1) + " of " + worksheet.Rows.Length;
                    Application.DoEvents();

                    writeCell = 0;
                    int k = 0;               

                    foreach (var cell in row.Cells)
                    {
                        writeCell=dataTransfer.testIfDoneExcel(j, k, cell, writeCell);
                        dataTransfer.transferExcelData(j, k, i, cell, form1, writeCell);
                        k++;
                    }
                    bool spectraExists = dataTransfer.writeSpectra(form1);
                    if (spectraExists == true) { dataTransfer.createConfirmFile(form1.location); }
                    form1.name = null;
                    form1.location = null;
                    dataTransfer.fileExists = false;
                    j++;
                }                
            }
        }
    }
}

