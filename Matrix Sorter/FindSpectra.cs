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


        internal static bool readWriteSpectra(string location, Form1 formObject)
        {
            Form1 form1 = formObject;
            string input = "empty";
            bool spectraFound = true;
            try
            {
                input = File.ReadAllText(string.Concat(location, ".txt"));
            }
            catch (FileNotFoundException e)
            {
                spectraFound = false;
                form1.missingSpectra.Add(form1.name);
            }

            if (spectraFound == true)
            {
                int pos = 0;
                int stop;
                int a = 0;

                while (a == 0)
                {
                    pos = input.IndexOf("\t", pos) + 1;
                    stop = input.IndexOf("\r\n", pos);
                    string val = input.Substring(pos, stop - pos);

                    using (StreamWriter sw = File.AppendText(form1.file))
                    {
                        sw.Write(string.Concat(val, ";"));

                        sw.Close();
                    }
                    if (pos > input.Length - 15)
                    {
                        a = 1;
                    }
                }
            }
            return spectraFound;

        }
    }
}
