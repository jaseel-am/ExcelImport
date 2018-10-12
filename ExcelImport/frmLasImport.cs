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

using ClosedXML.Excel;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelImport
{
    public partial class frmLasImport : Form
    {
        Excel.Workbook xlWorkBook;
        Excel.Application xlApp;
        Excel.Worksheet xlWorkSheet;
        Excel.Range Startrange;
        public frmLasImport()
        {
            InitializeComponent();
        }
        private void ReadFromTxtFile(string filePath, bool isFirst, bool isLast, int inOrder)
        {
            try
            {
                string lasData = System.IO.File.ReadAllText(filePath);
                int inAscciIndex = 0;
                int inLogIndex = 0;
                int inWellInfoIndex = 0;
                string[] lines = System.IO.File.ReadAllLines(filePath);

                for (int i = 0; i < lines.Length; i++)
                {
                    string strLine = lines[i];
                    if (strLine == "~Ascii")
                    {
                        inAscciIndex = i;
                    }
                    else if (strLine == "~Curve Information")
                    {
                        inLogIndex = i + 5;
                    }
                    else if (strLine == "~Well Information")
                    {
                        inWellInfoIndex = i + 5;
                    }
                }
                CreateTable(lines, inLogIndex, inAscciIndex, inWellInfoIndex, isFirst, isLast, inOrder);
                int s = inAscciIndex;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void CreateTable(string[] lines, int startIndex, int inAscciIndex, int inWellInfoIndex, bool isFirst, bool isLast, int inOrder)
        {
            try
            {
                DataTable dtbl = new DataTable();
                // column Name gets here
                for (int i = startIndex; i < inAscciIndex; i++)
                {
                    string ln = lines[i];
                    string[] LineData = ln.Split(' ');
                    string strColName = LineData[0];
                    dtbl.Columns.Add(strColName);
                }
                // column Datas get here
                inAscciIndex++;
                for (int i = inAscciIndex; i < lines.Length; i++)
                {
                    string ln = lines[i];
                    string[] LineData = ln.Split(' ');
                    LineData = LineData.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                    if (LineData.Length == dtbl.Columns.Count)
                    {
                        DataRow dr = dtbl.NewRow();
                        for (int j = 0; j < LineData.Length; j++)
                        {
                            dr[j] = LineData[j];
                        }
                        dtbl.Rows.Add(dr);
                    }
                    else
                    {
                        MessageBox.Show("Column numbers Dose not match", "EPPMS-Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                    }
                }
                // well Name and X Y values get here
                string strWellName = string.Empty;
                decimal decXVal = 0;
                decimal decYVal = 0;
                for (int i = inWellInfoIndex; i < lines.Length; i++)
                {
                    string ln = lines[i];
                    string[] LineData = ln.Split(' ');
                    LineData = LineData.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                    if (LineData[0].ToUpper() == "SHWN")
                    {
                        strWellName = LineData[2];
                    }
                    else if (LineData[0].ToUpper() == "ORIGINALWELLNAME")
                    {
                        strWellName = LineData[2];
                    }
                    else if (LineData[0].ToUpper() == "X")
                    {
                        decXVal = Convert.ToDecimal(LineData[2]);
                    }
                    else if (LineData[0].ToUpper() == "Y")
                    {
                        decYVal = Convert.ToDecimal(LineData[2]);
                        break;
                    }
                }

                DataColumn Col = dtbl.Columns.Add("WellName");
                Col.SetOrdinal(0);
                dtbl.Columns["WellName"].Expression = "'" + strWellName + "'";

                dtbl.Columns.Add("X");
                dtbl.Columns.Add("Y");

                dtbl.Columns["X"].Expression = "'" + decXVal + "'";
                dtbl.Columns["Y"].Expression = "'" + decYVal + "'";

                int rowCount = dtbl.Rows.Count;
                if (dtbl.Rows.Count > 0)
                {
                    //dataGridView1.DataSource = dtbl;
                    string[] columns = new string[] { "WellName", "MD", "CALI", "DEN", "GR", "NEU", "RES_DEP", "RES_SLW", "TVDSS", "X", "Y" };
                    if (dtbl.Columns.Contains("PEF") && dtbl.Columns.Contains("TVDSS"))
                    {
                        columns = new string[] { "WellName", "MD", "CALI", "DEN", "PEF", "GR", "NEU", "RES_DEP", "RES_SLW", "TVDSS", "X", "Y" };
                    }
                    else if (dtbl.Columns.Contains("DEPT"))
                    {
                        columns = new string[] { "WellName", "DEPT", "CALI", "DEN", "PEF", "GR", "NEU", "RES_DEP", "RES_SLW" };
                    }
                    DataTable newTable = dtbl.DefaultView.ToTable(false, columns);
                    // DataTable newTable = dtbl.DefaultView.ToTable(false, "WellName", "MD", "CALI", "DEN", "GR", "NEU", "RES_DEP", "RES_SLW", "TVDSS", "X", "Y");
                    dataGridView1.DataSource = newTable;
                    GenerateExcel(newTable, isFirst, isLast, inOrder);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void GenerateExcel(DataTable dtbl, bool isFirst, bool isLast, int SectionOrder)
        {
            try
            {
                bool isResult = StartExport(dtbl, isFirst, isLast, txtOutPutPath.Text.Trim(), txtTemplatePath.Text, "EPMS.xlsx", SectionOrder);
                if (isResult)
                {
                    MessageBox.Show("Processed Successfully..! Please check the output folder.!", "EPMS-Success", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

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

        private int GetOrder(string strFileName)
        {
            int inOrder = 0;
            try
            {
                switch (strFileName)
                {
                    case
                        "MARMUL-55H1":
                        inOrder = 1;
                        break;
                    case
                        "MARMUL-100H1":
                        inOrder = 2;
                        break;

                    case
                        "MARMUL-173H2":
                        inOrder = 3;
                        break;
                    case
                        "MARMUL-273H2":
                        inOrder = 4;
                        break;
                    case
                        "MARMUL-848H1":
                        inOrder = 5;
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {

                throw;
            }
            return inOrder;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Las Files|*.las";
                DialogResult result = ofd.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtFilePath.Text = ofd.FileName;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private void btnBrowseFolder_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtFilePath.Text != string.Empty)
                {
                    int inCount = 1;
                    string strTex = txtFilePath.Text.Trim();
                    if (strTex.Contains(".las"))
                    {
                        string strFileName = Path.GetFileNameWithoutExtension(txtFilePath.Text);
                        inCount = GetOrder(strFileName);
                        ReadFromTxtFile(txtFilePath.Text, true, true, inCount);
                    }
                    else
                    {
                        var fileCount = (from file in Directory.EnumerateFiles(@strTex, "*.las", SearchOption.AllDirectories)
                                         select file).Count();

                        bool isFirst = true; bool isLast = false;
                        foreach (string file in Directory.EnumerateFiles(strTex, "*.las"))
                        {
                            string strFileName = Path.GetFileNameWithoutExtension(file);
                            inCount = GetOrder(strFileName);
                            //isFirst = inCount == 1 ? true : false;
                            isLast = fileCount == inCount ? true : false;
                            ReadFromTxtFile(file, isFirst, isLast, inCount);
                            inCount++;
                            isFirst = false;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Select a file first..!", "EPPMS-Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "EPPMS-Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK) // Test result.
                {
                    txtFilePath.Text = folderBrowserDialog1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void frmLasImport_Load(object sender, EventArgs e)
        {

        }
    }
}
