using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matrix_Sorter
{
    class Matrix_Writer
    {
        
        internal static void writeCell(string text, Form1 formObject)
        {
            Form1 form1 = formObject;


            using (StreamWriter sw2 = File.AppendText(form1.file))
            {
                
                    sw2.Write(string.Concat(text, ";"));
                
                sw2.Close();                           
            }
        }


    }
}
