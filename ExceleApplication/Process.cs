using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace ExceleApplication
{
    public class Process
    {
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;


        public void ReadExistingExcel(List<ExcellData> data)
        {
            string path = @"D:\Users\udagasan\Source\Repos\MountainExceeder\ExceleApplication\Sources\İSO-DER-SM-MEMUR (mdm) 17.11.Ay.xls";
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            mWorkSheets = mWorkBook.Worksheets;
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("kurummaas");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;

            int rowBeginning = 10;

            foreach (var item in data)
            {
                rowBeginning++;

                for (int index = 1; index < 2; index++)
                {
                    mWSheet1.Cells[rowBeginning, 1] = item.FullName;
                    mWSheet1.Cells[rowBeginning, 2] = item.AccountNumber;
                    mWSheet1.Cells[rowBeginning, 3] = item.RegisterNumber;
                    mWSheet1.Cells[rowBeginning, 4] = item.Amount;
                    mWSheet1.Cells[rowBeginning, 5] = item.Iban;
                }
            }
            var day = DateTime.Now.Day.ToString();
            var month = DateTime.Now.Month.ToString();
            var year = DateTime.Now.Year.ToString().Substring(2, 2);

            string name = string.Concat("İSO-DER-SM-MEMUR (mdm) {0}.{1}.{2}.xsl",year,day,month);
            string destinatonPath = string.Concat(@"D:\Users\udagasan\Source\Repos\MountainExceeder\ExceleApplication\Sources\",name);

            mWorkBook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void CreateNewExcellAndFillFromDataTable()
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true,
                    DisplayAlerts = false
                };
                worKbooK = excel.Workbooks.Add(Type.Missing);

                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "StudentRepoertCard";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "Student Report Card";
                worKsheeT.Cells.Font.Size = 15;


                int rowcount = 2;

                foreach (DataRow datarow in GetData().Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= GetData().Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = GetData().Columns[i - 1].ColumnName;
                            worKsheeT.Cells.Font.Color = ConsoleColor.Black;

                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == GetData().Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, GetData().Columns.Count]];
                                }

                            }
                        }

                    }

                }

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, GetData().Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, GetData().Columns.Count]];

                worKbooK.SaveAs(@"D:\Users\udagasan\Desktop\A");
                //var a = File.CreateText(@"D:\Users\udagasan\Source\Repos\MountainExceeder\ExceleApplication\Sources");
                worKbooK.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
                throw new Exception(ex.Message);

            }
            finally
            {
                worKsheeT = null;
                celLrangE = null;
                worKbooK = null;
            }
        }

        private DataTable GetData()
        {
            throw new NotImplementedException();
        }
    }



    public class ExcellData
    {
        public string FullName;
        public string AccountNumber;
        public string RegisterNumber;
        public string Amount;
        public string Currenc;
        public string Iban;

    }
}
