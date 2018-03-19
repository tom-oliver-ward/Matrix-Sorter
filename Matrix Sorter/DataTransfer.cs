using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matrix_Sorter
{
    class DataTransfer
    {
        public bool fileExists = false;
        internal int testIfDoneExcel(int j, int k, Excel.Cell cell,int writeCell)
        {
            
            if (j > 0 && k == 0 && cell == null)
            {
                writeCell = 1;
            }
            return writeCell;
        }

        internal void transferExcelData(int j, int k,int i, Excel.Cell cell, Form1 formObject, int writeCell)
        {
            Form1 form1 = formObject;
            if (j > 0 && k != 0 && writeCell == 1)
            {
                
                if (k == 1 && cell!=null) 
                {
                    form1.name = cell.Text;
                    LocationSet(form1, i);
                }



                if (form1.name != null && form1.location != null && k == 1)
                {
                    fileExists = TestIfDoneText(form1.location);
                }

                if (cell != null && fileExists == false && form1.location != null)
                {
                    Matrix_Writer.writeCell(cell.Text, form1);
                }
                else if (k > 1 && form1.location != null && fileExists == false)
                {
                    Matrix_Writer.writeCell(" ", form1);
                }
            }
        }

        private void LocationSet(Form1 formObject, int i)
        {
            Form1 form1 = formObject;
            form1.location = (string)form1.SpreadSheets2Convert.Items[i];
            int pos=1;
            int posStore=0;

            while (pos>=1)
            {
                posStore = pos;
                pos = form1.location.IndexOf("\\", pos) + 1;

            }
            //Find last point in intList, take all of location prior to this point
            form1.location = form1.location.Substring(0, posStore);

            form1.location = string.Concat(form1.location, "Processed\\", form1.name);

        }

        private bool TestIfDoneText(string location)
        {
            bool value = File.Exists(string.Concat(location, "Sorted.txt"));
            return value;
        }

        internal bool writeSpectra(Form1 formObject)
        {
            Form1 form1 = formObject;
            bool spectraFound = false;
            if (form1.location != null && fileExists == false)
            {
                spectraFound = FindSpectra.readWriteSpectra(form1.location, form1);
                newline(form1);                
            }
            return spectraFound;            
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

        internal void createConfirmFile(string location)
        {
            if (location!=null)
            {
                location=(string.Concat(location, "Sorted.txt"));
                using (File.Create(location));
            }
            
        }
    }
}
