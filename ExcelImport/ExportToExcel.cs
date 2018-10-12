using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using System.Windows.Forms;
using System.IO;

namespace ExcelImport
{
    class ExportToExcel
    {
        Excel.Workbook xlWorkBook;
        Excel.Application xlApp;
        Excel.Worksheet xlWorkSheet;
        Excel.Range Startrange;

        public bool StartExport(DataTable dtbl, bool isFirst, bool isLast, string strOutputPath, string TemplateLocation, string TemplateFullName, int SectionOrder)
        {
            bool isSuccess = false;
            try
            {
                if (isFirst)
                {
                    CopyTemplate(TemplateLocation, strOutputPath, TemplateFullName);
                    xlApp = new Excel.Application();
                    if (xlApp == null)
                    {
                        throw new Exception("Excel is not properly installed!!");
                    }
                    xlWorkBook = xlApp.Workbooks.Open(@strOutputPath + TemplateFullName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                }
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(SectionOrder);
                Startrange = xlWorkSheet.get_Range("A2");
                FillInExcel(Startrange, dtbl);
                if (isLast)
                {
                    xlApp.DisplayAlerts = false;
                    xlWorkBook.SaveAs(@strOutputPath + TemplateFullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
                    xlWorkBook.Close(true, null, null);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    isSuccess = true;
                }
            }
            catch (Exception ex)
            {
                isSuccess = false;
                throw;
            }
            return isSuccess;
        }
        private void FillInExcel(Excel.Range startrange, DataTable dtblData)
        {
            int rw = 0;
            int cl = 0;
            try
            {
                // Fill The Report Content Data Here
                rw = dtblData.Rows.Count;
                cl = dtblData.Columns.Count;
                string[,] data = new string[rw, cl];
                for (var row = 1; row <= rw; row++)
                {
                    for (var column = 1; column <= cl; column++)
                    {
                        data[row - 1, column - 1] = dtblData.Rows[row - 1][column - 1].ToString();
                    }
                }
                Excel.Range endRange = (Excel.Range)xlWorkSheet.Cells[rw + (startrange.Cells.Row - 1), cl + (startrange.Cells.Column - 1)];
                Excel.Range writeRange = xlWorkSheet.Range[startrange, endRange];
                writeRange.WrapText = true;
                writeRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                writeRange.Value2 = data;
                writeRange.Formula = writeRange.Formula;
                data = null;
                startrange = null;
                endRange = null;
                writeRange = null;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void CopyTemplate(string strFrom, string strTo, string FileName)
        {
            try
            {
                if (!Directory.Exists(strTo))
                {
                    Directory.CreateDirectory(strTo);
                }
                if (File.Exists(strFrom + FileName))
                {
                    if (!File.Exists(strTo + FileName))
                    {
                        File.Copy(strFrom + FileName, strTo + FileName, true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
