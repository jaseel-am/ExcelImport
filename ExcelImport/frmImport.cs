using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelImport
{
    public partial class frmImport : Form
    {
        frmLoading frmLodingObj = new frmLoading();
        public frmImport()
        {
            InitializeComponent();
        }
        private DataTable ReadExcelFile(string sheetName, string path, string strMode)
        {
            DataTable dt = new DataTable();
            try
            {
                using (OleDbConnection conn = new OleDbConnection())
                {
                    string Import_FileName = path;
                    string fileExtension = Path.GetExtension(Import_FileName);
                    if (fileExtension == ".xls")
                        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                    if (fileExtension == ".xlsx")
                        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                    using (OleDbCommand comm = new OleDbCommand())
                    {
                        if (strMode == "First")
                        {
                            comm.CommandText = "Select * from [" + sheetName + "$A1:V500000]";
                        }
                        else
                        {
                            comm.CommandText = "Select * from [" + sheetName + "$X1:AS500000]";
                        }
                        comm.Connection = conn;
                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            da.SelectCommand = comm;
                            da.Fill(dt);
                            return dt;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OKCancel);
            }
            return dt;
        }
        private void btnProcess_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtInput.Text != string.Empty)
                {
                    bwrk1.RunWorkerAsync();
                    frmLodingObj.ShowFromOtherForms();
                }
                else
                {
                    MessageBox.Show("Select the file first", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }
        private void copyAlltoClipboard()
        {
            DgvResult.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DgvResult.MultiSelect = true;
            DgvResult.SelectAll();
            DataObject dataObj = DgvResult.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private DataTable StartProcess()
        {
            DataTable dtblOutput = new DataTable();
            try
            {
                DataTable dtblFirst = ReadExcelFile("Sheet1", txtInput.Text.Trim(), "First");
                DataTable dtblSecond = ReadExcelFile("Sheet1", txtInput.Text.Trim(), "Second");
                dtblOutput = dtblSecond.Clone();
                if (!dtblOutput.Columns.Contains("Distance"))
                {
                    dtblOutput.Columns.Add("Distance");
                }
                if (dtblFirst.Rows.Count > 0)
                {
                    if (dtblFirst.Columns.Contains("WELL") && dtblFirst.Columns.Contains("DEPTH") && dtblFirst.Columns.Contains("Abs-X(meters)") && dtblFirst.Columns.Contains("Abs-Y(meters)"))
                    {
                        if (dtblSecond.Columns.Contains("WELL") && dtblSecond.Columns.Contains("DEPTH") && dtblSecond.Columns.Contains("Abs-X(meters)") && dtblSecond.Columns.Contains("Abs-Y(meters)"))
                        {
                            for (int i = 0; i < dtblFirst.Rows.Count; i++)
                            {
                                if (dtblFirst.Rows[i]["WELL"].ToString() != string.Empty)
                                {
                                    double FDepth = Convert.ToDouble(dtblFirst.Rows[i]["DEPTH"].ToString());
                                    double FAbsXMeter = Convert.ToDouble(dtblFirst.Rows[i]["Abs-X(meters)"].ToString());
                                    double FbsYmeters = Convert.ToDouble(dtblFirst.Rows[i]["Abs-Y(meters)"].ToString());

                                    double SDepth = 0;
                                    double SAbsXMeter = 0;
                                    double SbsYmeters = 0;

                                    for (int j = 0; j < dtblSecond.Rows.Count; j++)
                                    {
                                        if (dtblFirst.Rows[i]["WELL"].ToString() != string.Empty)
                                        {
                                            SDepth = Convert.ToDouble(dtblSecond.Rows[j]["DEPTH"].ToString());
                                            SAbsXMeter = Convert.ToDouble(dtblSecond.Rows[j]["Abs-X(meters)"].ToString());
                                            SbsYmeters = Convert.ToDouble(dtblSecond.Rows[j]["Abs-Y(meters)"].ToString());

                                            double decDepthSum = FDepth - SDepth;
                                            double decAbsXMeterSum = FAbsXMeter - SAbsXMeter;
                                            double decAbsYMeterSum = FbsYmeters - SbsYmeters;

                                            double decDepthSqure = decDepthSum * decDepthSum;
                                            double decAbsXMeterSqure = decAbsXMeterSum * decAbsXMeterSum;
                                            double decAbsYMeterSqure = decAbsYMeterSum * decAbsYMeterSum;
                                            double grandSum = decDepthSqure + decAbsXMeterSqure + decAbsYMeterSqure;

                                            double LastAns = Math.Sqrt(grandSum);
                                            if (!dtblSecond.Columns.Contains("Distance"))
                                            {
                                                dtblSecond.Columns.Add("Distance");
                                            }
                                            dtblSecond.Rows[j]["Distance"] = LastAns;
                                        }
                                    }
                                    double minAccountLevel = double.MaxValue;
                                    double maxAccountLevel = double.MinValue;
                                    for (int k = 0; k < dtblSecond.Rows.Count; k++)
                                    {
                                        double accountLevel = Convert.ToDouble(dtblSecond.Rows[k]["Distance"].ToString());
                                        minAccountLevel = Math.Min(minAccountLevel, accountLevel);
                                        maxAccountLevel = Math.Max(maxAccountLevel, accountLevel);
                                    }
                                    DataRow dr;
                                    for (int l = 0; l < dtblSecond.Rows.Count; l++)
                                    {
                                        if (Convert.ToDouble(dtblSecond.Rows[l]["Distance"].ToString()) == minAccountLevel)
                                        {
                                            dr = dtblSecond.Rows[l];
                                            dtblOutput.Rows.Add(dr.ItemArray);
                                        }
                                    }
                                    string s = string.Empty;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("The Uploaded Excel file column not matching", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("The Uploaded Excel file column not matching", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No data found in Excel file", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
            return dtblOutput;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (DgvResult.Rows.Count > 0)
                {
                    copyAlltoClipboard();
                    Microsoft.Office.Interop.Excel.Application xlexcel;
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlexcel = new Microsoft.Office.Interop.Excel.Application();
                    xlexcel.Visible = true;
                    xlWorkBook = xlexcel.Workbooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                    CR.Select();
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                }
                else
                {
                    MessageBox.Show("Process the file first", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsb";
                DialogResult result = ofd.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtInput.Text = ofd.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }

        private void frmImport_Load(object sender, EventArgs e)
        {
            try
            {
                txtInput.Text = string.Empty;
                txtInput.ReadOnly = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }

        private void bwrk1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                DataTable dtbl = StartProcess();
                Thread thread = new Thread(() => SetDataSource(dtbl));
                thread.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void bwrk1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (frmLodingObj != null)
                {
                    frmLodingObj.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        delegate void SetDataSourceCallback(DataTable dtblOutput);

        public void SetDataSource(DataTable dtMessage)
        {
            try
            {
                if (this.DgvResult.InvokeRequired)
                {
                    SetDataSourceCallback d = new SetDataSourceCallback(SetDataSource);
                    this.Invoke(d, new object[] { dtMessage });
                }
                else
                {
                    DgvResult.DataSource = dtMessage;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SetDataSource " + ex.ToString());
            }
        }
    }
}
