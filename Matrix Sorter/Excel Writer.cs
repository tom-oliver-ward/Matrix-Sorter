using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace Matrix_Sorter
{
    class Excel_Writer
    {

        private string excelFilePath = string.Empty;
        private int rowNumber = 1; // define first row number to enter data in excel

        MyExcel.Application myExcelApplication;
        MyExcel.Workbook myExcelWorkbook;
        MyExcel.Worksheet myExcelWorkSheet;

        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }

        public int Rownumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }

        public void openExcel(string filename)
        {
            myExcelApplication = null;

            myExcelApplication = new MyExcel.Application(); // create Excell App
            myExcelApplication.DisplayAlerts = false; // turn off alerts


            myExcelWorkbook = (MyExcel.Workbook)(myExcelApplication.Workbooks._Open(filename, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

            int numberOfWorkbooks = myExcelApplication.Workbooks.Count; // get number of workbooks (optional)

            myExcelWorkSheet = (MyExcel.Worksheet)myExcelWorkbook.Worksheets[1]; // define in which worksheet, do you want to add data
            myExcelWorkSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

            int numberOfSheets = myExcelWorkbook.Worksheets.Count; // get number of worksheets (optional)
        }



            public void addDataToExcel(string firstname, string lastname, string language, string email, string company)
        {

            myExcelWorkSheet.Cells[rowNumber, "H"] = firstname;
            myExcelWorkSheet.Cells[rowNumber, "J"] = lastname;
            myExcelWorkSheet.Cells[rowNumber, "Q"] = language;
            myExcelWorkSheet.Cells[rowNumber, "BH"] = email;
            myExcelWorkSheet.Cells[rowNumber, "CH"] = company;
            rowNumber++;  // if you put this method inside a loop, you should increase rownumber by one or wat ever is your logic

        }

        public void closeExcel(string filename)
        {
            try
            {
                myExcelWorkbook.SaveAs(filename, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, MyExcel.XlSaveAsAccessMode.xlNoChange,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value); // Save data in excel


                myExcelWorkbook.Close(true, filename, System.Reflection.Missing.Value); // close the worksheet


            }
            finally
            {
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                }
            }

        }
    

    internal void writeConfirm(Form1 formObject, int j, int i)
        {

            Form1 form1 = formObject;
            string filename = (string)form1.SpreadSheets2Convert.Items[i];
            string confirm="1";

            openExcel(filename);
            addDataToExcel("a", "a", "a", "a", "a");
            closeExcel(filename);



            //Microsoft.Office.Interop.Excel.Application oXL;
            //Microsoft.Office.Interop.Excel._Workbook oWB;
            //Microsoft.Office.Interop.Excel._Worksheet oSheet;
            //Microsoft.Office.Interop.Excel.Range oRng;
            //object misvalue = System.Reflection.Missing.Value;
            
            //    oXL = new Microsoft.Office.Interop.Excel.Application();
            //   // oXL.Visible = true;

            
            //oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(filename));
            //oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            //oSheet.Cells[1, j] = confirm;

            //oWB.Save();
            //oWB.SaveAs("C:\\Users\tom2s\\Google Drive\\PF\\Orbis\\software\\File Organizer Matrix\\Example\\test.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            //false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            //Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //oWB.Close();
            


            // Get the Excel application object.
            //MyExcel.Application excel_app = new MyExcel.Application();

            //// Make Excel visible (optional).
            ////excel_app.Visible = true;



            //// Open the workbook.
            //MyExcel.Workbook workbook = excel_app.Workbooks.Open(filename,
            //    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //    Type.Missing, Type.Missing);

            //// See if the worksheet already exists.
            //string sheet_name = "Sheet1";

            //MyExcel.Worksheet worksheet = (MyExcel.Worksheet)workbook.Worksheets[sheet_name];
            //worksheet.Cells[1, j] = confirm;

            //// Save the changes and close the workbook.
            //workbook.Close(true, Type.Missing, Type.Missing);

            //// Close the Excel server.
            //excel_app.Quit();

        }


    }


}
