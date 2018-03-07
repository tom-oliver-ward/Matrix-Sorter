using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matrix_Sorter
{
    class FindSpectra
    {


        internal static void readWriteSpectra(string location, Form1 formObject)
        {
            Form1 form1 = formObject;
            string input = File.ReadAllText(string.Concat(location,".txt"));
            int pos=0;            
            int stop;            
            int a = 0;                 

            while (a==0)
            {
                pos = input.IndexOf("\t",pos) + 1;
                stop = input.IndexOf("\r\n", pos);
                string val = input.Substring(pos, stop - pos);

                using (StreamWriter sw = File.AppendText(form1.file))
                {
                    sw.Write(string.Concat(val, ";"));                
                    
                    sw.Close();
                }
                if (pos>input.Length-15)
                {
                    a = 1;
                }
            }
        }
    }
}
