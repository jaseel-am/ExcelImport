using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelImport
{
    public class InteropExcelExport
    {
        Excel.Workbook xlWorkBook;
        Excel.Application xlApp;
        Excel.Worksheet xlWorkSheet;
        Excel.Range Startrange;
        Excel.Range HeaderStartrange;
        public bool StartExport(DataTable dtbl, bool isFirst, bool isLast, string strOutputPath, string TemplateLocation, string TemplateFullName, int SectionOrder, int totalNoOfSheets)
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

                    // To Add Sheets Dynamically
                    for (int i = 0; i <= totalNoOfSheets; i++)
                    {
                        int count = xlWorkBook.Worksheets.Count;
                        Excel.Worksheet addedSheet = xlWorkBook.Worksheets.Add(Type.Missing,
                                xlWorkBook.Worksheets[count], Type.Missing, Type.Missing);
                        addedSheet.Name = "Sheet " + i.ToString();
                    }
                }
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(SectionOrder);
                Startrange = xlWorkSheet.get_Range("A2");
                HeaderStartrange = xlWorkSheet.get_Range("A1");
                FillInExcel(Startrange, HeaderStartrange, dtbl);
                xlWorkSheet.Name = dtbl.TableName;
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
                MessageBox.Show(ex.Message, "LEXEL: ERROR ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return isSuccess;
        }

        private void FillInExcel(Excel.Range startrange, Excel.Range HeaderStartRange, DataTable dtblData)
        {
            int rw = 0;
            int cl = 0;
            try
            {
                // Fill The Report Content Data Here
                rw = dtblData.Rows.Count;
                cl = dtblData.Columns.Count;
                string[,] data = new string[rw, cl];

                // Adding Data Table Header Here
                for (int i = 1; i <= dtblData.Columns.Count; i++)
                {
                    string strName = dtblData.Columns[i - 1].ColumnName;
                    (HeaderStartRange.Cells[1, i] as Excel.Range).Value2 = strName;
                }

                // Adding Columns Here
                for (var row = 1; row <= rw; row++)
                {
                    for (var column = 1; column <= cl; column++)
                    {
                        data[row - 1, column - 1] = dtblData.Rows[row - 1][column - 1].ToString();
                    }
                }
                Excel.Range endRange = (Excel.Range)xlWorkSheet.Cells[rw + (startrange.Cells.Row - 1), cl + (startrange.Cells.Column - 1)];
                Excel.Range writeRange = xlWorkSheet.Range[startrange, endRange];
                //writeRange.WrapText = true;
                //writeRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                writeRange.Value2 = data;
                writeRange.Formula = writeRange.Formula;
                data = null;
                startrange = null;
                endRange = null;
                writeRange = null;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "LEXEL: ERROR ", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    else
                    {
                        File.Delete(strTo + FileName);
                        File.Copy(strFrom + FileName, strTo + FileName, true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "LEXEL: ERROR ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
