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

        internal string transferExcelData(int j, int k,int i, Excel.Cell cell, Form1 formObject, int writeCell, string name, string location)
        {
            Form1 form1 = formObject;
            if (j > 0 && k != 0 && writeCell == 1)
            {
                
                if (k == 1 && cell!=null) 
                {
                    name = cell.Text;
                    location = LocationSet(form1, i, name, location);
                }
                
                               

                if (name != null && location != null && k==1)
                {
                    fileExists = TestIfDoneText(location);
                }

                if (cell != null && fileExists == false && location != null)
                {
                    Matrix_Writer.writeCell(cell.Text, form1);
                }
                else if (k > 1 && location != null && fileExists == false)
                {
                    Matrix_Writer.writeCell(" ", form1);
                }
            }

            return location;
        }

        private string LocationSet(Form1 formObject, int i, string name, string location)
        {
            Form1 form1 = formObject;
            location = (string)form1.SpreadSheets2Convert.Items[i];
            int pos=1;
            int posStore=0;

            while (pos>=1)
            {
                posStore = pos;
                pos = location.IndexOf("\\",pos) + 1;

            }
            //Find last point in intList, take all of location prior to this point
            location = location.Substring(0, posStore);

            location = string.Concat(location, "Processed\\", name);
            return location;
        }

        private bool TestIfDoneText(string location)
        {
            bool value = File.Exists(string.Concat(location, "Sorted.txt"));
            return value;
        }

        internal void writeSpectra(Form1 formObject, string name, string location)
        {
            Form1 form1 = formObject;

            if (location != null && fileExists==false)
            {
                FindSpectra.readWriteSpectra(location, form1);
                newline(form1);                
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
