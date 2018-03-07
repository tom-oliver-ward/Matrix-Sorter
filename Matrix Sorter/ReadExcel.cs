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
        Excel_Writer excelWriter = new Excel_Writer();

        public void readFile(int i, Form1 formObject)
        {
            Form1 form1 = formObject;
            string filename = (string)form1.SpreadSheets2Convert.Items[i];


            int j = 0;
            
            foreach (var worksheet in Workbook.Worksheets(filename))
            {
                foreach (var row in worksheet.Rows)
                {
                    int k = 0;
                    int writeCell = 0;
                    string name=null;
                    string location=null;
                    foreach (var cell in row.Cells)
                    {
                        if (j>0 && k==0 && cell==null)
                        {
                            writeCell = 1;
                            excelWriter.writeConfirm(form1, j, i);                            
                        }

                        if (j > 0 && k!=0 && writeCell==1)
                        {
                            if (k == 1) { name = cell.Text; }
                            if (k == 2) { location = cell.Text; }

                            if (cell != null) { Matrix_Writer.writeCell(cell.Text, form1); }
                            else { Matrix_Writer.writeCell(" ", form1); }                            
                        }           


                        k++;
                    }
                    //if (j>0)
                    //{
                    //    Matrix_Writer.newline(form1);
                    //}
                    j++;
                    if (name!=null)
                    { FindSpectra.readSpectra(name, location, form1);

                        using (StreamWriter sw3 = File.AppendText(form1.file))
                        {

                            sw3.WriteLine(string.Concat(" "));

                            sw3.Close();
                        }
                    }


                }

                
            }

        }



    }
}

