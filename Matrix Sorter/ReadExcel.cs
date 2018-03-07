using Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace Matrix_Sorter
{
    class ReadExcel
    {
        DataTransfer dataTransfer = new DataTransfer();
        Matrix_Writer matrixWriter = new Matrix_Writer();
        
             
        public int writeCell;
        public string name = null;
        public string location = null;

        public void readFile(int i, Form1 formObject)
        {
            Form1 form1 = formObject;
            string filename = (string)form1.SpreadSheets2Convert.Items[i];
            int j = 0;
            
            foreach (var worksheet in Workbook.Worksheets(filename))
            {
                foreach (var row in worksheet.Rows)
                {
                    writeCell = 0;
                    int k = 0;               

                    foreach (var cell in row.Cells)
                    {
                        writeCell=dataTransfer.testIfDoneExcel(j, k, cell, writeCell);
                        location = dataTransfer.transferExcelData(j, k, i, cell, form1, writeCell, name, location);
                        k++;
                    }                    
                    dataTransfer.writeSpectra(form1, name, location);
                    j++;


                }

                
            }

        }



    }
}

