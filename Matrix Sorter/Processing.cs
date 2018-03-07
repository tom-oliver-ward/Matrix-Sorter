using System;
using System.Collections.Generic;
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
            }
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show("Complete", "", buttons);
        }


    }
}
