using System;
using System.Collections.Generic;
using System.Diagnostics;
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

                
                if (form1.missingSpectra.Count > 0)
                {
                    FailedSpectra(form1, i);
                    //Process.Start(form1.fileError);
                }
                form1.missingSpectra.Clear();
            }
            
        }



        internal void FailedSpectra(Form1 formObject,int i)
        {
            Form1 form1 = formObject;            
            using (StreamWriter sw2 = File.AppendText(string.Concat(form1.fileError)))
            {
                sw2.WriteLine((string)form1.SpreadSheets2Convert.Items[i]);  
                for(int j=0; j<form1.missingSpectra.Count;j++)
                {
                    sw2.WriteLine(form1.missingSpectra[j]);
                }   

                sw2.Close();
            }
        }

        internal void FileErrorPath(Form1 formObject)
        {
            Form1 form1 = formObject;
            string removeTxt=form1.file.Substring(0, form1.file.Length - 4);
            form1.fileError = string.Concat(removeTxt, "Errors.txt");
        }
    }
}
