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

        Matrix_Writer matrixWriter = new Matrix_Writer();
        Processing processing = new Processing();        
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
                        processing.testIfDoneExcel(j, k, cell);
                        processing.transferExcelData(j, k, cell, form1); 
                        k++;
                    }
                    //if (j>0)
                    //{
                    //    Matrix_Writer.newline(form1);
                    //}
                    j++;
                    processing.writeSpectra(form1);                  


                }

                
            }

        }



    }
}

