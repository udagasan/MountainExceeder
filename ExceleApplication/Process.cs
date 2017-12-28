using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;


namespace ExceleApplication
{
    public class Process
    {
        private static Microsoft.Office.Interop.Excel.Workbook workBook;
        private static Microsoft.Office.Interop.Excel.Sheets workSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private static Microsoft.Office.Interop.Excel.Application excel;


        public string ReadExistingExcel(List<ExcellData> data, string currency, string day, string month, string year)
        {

           var path= GetFilePath("Memur");

            excel = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true,
                DisplayAlerts = false
            };
           
            try
            {
                workBook = excel.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            excel.Visible = false;
            workSheets = workBook.Worksheets;
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workSheets.get_Item("kurummaas");
            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;
            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;


            worksheet.Cells[7, 3] = month;
            worksheet.Cells[8, 4] = currency;
            worksheet.Cells[4, 3] = string.Concat(day, "/", month, "/", year);

            //worksheet.Cells[5, 4] = totalAmount;
            //worksheet.Cells[6, 4] = data.Count;
            int rowBeginning = 10;

            foreach (var item in data)
            {
                rowBeginning++;

                for (int index = 1; index < 2; index++)
                {
                    worksheet.Cells[rowBeginning, 1] = item.FullName;
                    worksheet.Cells[rowBeginning, 2] = item.AccountNumber;
                    worksheet.Cells[rowBeginning, 3] = item.RegisterNumber;
                    worksheet.Cells[rowBeginning, 4] = item.Amount;
                    worksheet.Cells[rowBeginning, 5] = item.Iban;
                }
            }
            year = year.Substring(2, 2);

            string name = string.Concat("\\Sources\\İSO-DER-SM-MEMUR (mdm) ", day, ".", month, ".", year, ".xls");
            string destinatonPath = string.Concat(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), name);
            workBook.SaveCopyAs(destinatonPath);
            workBook.Close(Missing.Value, Missing.Value, Missing.Value);
            worksheet = null;
            workBook = null;
            excel.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            return destinatonPath;
        }

        static string GetFilePath(string fileName)
        {

            var resourceNames = Assembly.GetExecutingAssembly().GetManifestResourceNames();
            string currentResource = "";
            foreach (var item in resourceNames)
            {
                if (item.Contains(fileName))
                {
                    currentResource = item;
                    break;
                }
            }

            var dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(dllPath, @"Sources\MemurTemp.xls");
            var file = Assembly.GetExecutingAssembly().GetManifestResourceStream(currentResource);

            return path;

        }

    }

    public class ExcellData
    {
        public string FullName;
        public string AccountNumber;
        public string RegisterNumber;
        public string Amount;
        public string Iban;

    }
}
