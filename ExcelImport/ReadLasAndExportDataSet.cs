
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Windows.Forms;

namespace ExcelImport
{
    class ReadLasAndExportDataSet
    {
        DataSet dsData = new DataSet();
        DataTable dtblCalcimetry = new DataTable();
        DataTable dtblDrillingParameter = new DataTable();
        DataTable dtblIntegrtdLithology = new DataTable();
        DataTable dtblRop = new DataTable();
        DataTable dtblHcIndicator = new DataTable();
        DataTable dtblLithologyPercntg = new DataTable();

        private void ReadFromTxtFile(string strInputFilePath)
        {
            try
            {
                int inColHdrstartIndex = 0;
                string lasData = System.IO.File.ReadAllText(strInputFilePath);
                int inAscciIndex = 0;
                int inColHeaderIndex = 0;
                string strUtme = string.Empty;
                string strUtmn = string.Empty;
                string[] lines = System.IO.File.ReadAllLines(strInputFilePath);
                for (int i = 0; i < lines.Length; i++)
                {
                    string strLine = lines[i];
                    if (strLine == "~ASCII")
                    {
                        inAscciIndex = i + 1;
                    }
                    else if (strLine.Contains("UTME."))
                    {
                        string ln = lines[i];
                        string[] LineData = ln.Split(' ');
                        LineData = LineData.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                        strUtme = LineData[1].TrimEnd(':');
                    }
                    else if (strLine.Contains("UTMN."))
                    {
                        string ln = lines[i];
                        string[] LineData = ln.Split(' ');
                        LineData = LineData.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                        strUtmn = LineData[1].TrimEnd(':');
                    }
                    else if (strLine == "~Curve Information Block")
                    {
                        inColHeaderIndex = i + 9;
                        inColHdrstartIndex = i + 3;
                    }
                }
                string strFullName = Path.GetFileNameWithoutExtension(strInputFilePath);
                string[] words = strFullName.Split('_');
                string strWellName = "";
                for (int i = 0; i < words.Length; i++)
                {
                    if (words[i].ToUpper() == "PUBLIC")
                    {
                        break;
                    }
                    else
                    {
                        strWellName += words[i] + "_";
                    }
                }
                CreateTable(lines, inColHdrstartIndex, inColHeaderIndex, strUtme, strUtmn, inAscciIndex, strWellName.TrimEnd('_'));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }
        private void CreateTable(string[] lines, int startIndex, int inColHeaderIndex, string strUtme, string strUtmn, int inAscciIndex, string strWellName)
        {
            try
            {
                DataTable dtbl = new DataTable();
                // column Name gets here
                for (int i = startIndex; i < inColHeaderIndex; i++)
                {
                    string ln = lines[i];
                    string[] LineData = ln.Split(' ');
                    string strColName = LineData[0];
                    dtbl.Columns.Add(strColName);
                }
                // column Datas get here
                inColHeaderIndex++;
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
                        Common.WriteToFile("Well: " + strWellName + " Column numbers Dose not match", false);
                    }
                }
                // well name and UTME columns adding here for entire data table.
                DataColumn Col = dtbl.Columns.Add("WellName");
                Col.SetOrdinal(0);
                dtbl.Columns["WellName"].Expression = "'" + strWellName + "'";
                dtbl.Columns.Add("UTME.");
                dtbl.Columns.Add("UTMN.");
                dtbl.Columns["UTME."].Expression = "'" + strUtme + "'";
                dtbl.Columns["UTMN."].Expression = "'" + strUtmn + "'";
                int rowCount = dtbl.Rows.Count;
                if (dtbl.Rows.Count > 0)
                {
                    dtbl.TableName = strWellName;

                    DataTable dtblCalcimetry = ReadExcelCalcimetry(strWellName.TrimEnd('_'));
                    DataTable dtblDrilling = ReadExcelDrillingParameters(strWellName.TrimEnd('_'));
                    DataTable dtblInterpLithlgy = ReadIntegretedLithology(strWellName.TrimEnd('_'));
                    DataTable dtbRop = ReadFromROP(strWellName.TrimEnd('_'));
                    DataTable dtblHcIndictr = ReadFromHCIndicator(strWellName.TrimEnd('_'));
                    DataTable dtblLithoPer = ReadFromLithologyPercentage(strWellName.TrimEnd('_'));


                    DataTable dt = MergeToFinalTable(strWellName, dtbl, dtblCalcimetry, dtblDrilling, dtblInterpLithlgy, dtbRop, dtblHcIndictr, dtblLithoPer);
                    if (!dsData.Tables.Contains(dt.TableName))
                    {
                        dsData.Tables.Add(dt);
                        Common.WriteToFile("Successfully Processed: " + strWellName, false);
                    }
                    else
                    {
                        Common.WriteToFile("Well details: " + strWellName + " Already added to the data set" + strWellName, false);
                    }   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }
        public bool ReadFileFromFolderandProcess(string strFolderpath, string strOutputPath, string TemplatePath, string strTemplateName)
        {
            bool isSuccess = false;
            try
            {
                int fileCount = (from file in Directory.EnumerateFiles(@strFolderpath, "*.las", SearchOption.AllDirectories)
                                 select file).Count();
                if (fileCount > 0)
                {
                    foreach (string file in Directory.EnumerateFiles(strFolderpath, "*.las"))
                    {
                        string strFileName = Path.GetFileNameWithoutExtension(file);
                        ReadFromTxtFile(file);
                    }
                    //   isSuccess = new ClosedXmlExport().ExportToExcelinClosedXml(dsData, strOutputPath); // Export Via Closed XML
                    isSuccess = ExportExcelInInterop(strOutputPath, TemplatePath, strTemplateName);
                    //ReadExcel(); // Calling here for temprory purpose
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("No files found in input location", "EPMS-Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return isSuccess;
        }
        private bool ExportExcelInInterop(string strOutputPath, string TemplatePath, string strTemplateName)
        {
            bool isSuccess = false;
            try
            {
                InteropExcelExport obj = new InteropExcelExport();
                bool isFirst = true;
                bool isLast = false;
                int inTotal = dsData.Tables.Count - 1;
                for (int i = 0; i < dsData.Tables.Count; i++)
                {
                    if (inTotal == i)
                    {
                        isLast = true;
                    }

                    isSuccess = obj.StartExport(dsData.Tables[i], isFirst, isLast, strOutputPath, TemplatePath, strTemplateName, i + 1, inTotal + 1);
                    isFirst = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return isSuccess;
        }

        private DataTable ReadExcelCalcimetry(string WellName)
        {
            DataTable dtbl = new DataTable(WellName);
            try
            {
                ReadFromExcel objClass = new ReadFromExcel();
                string strPath = "";
                string ExcelName = "Calcimetry";
                strPath = Properties.Settings.Default.CalcimetryPath;
                if (dtblCalcimetry.Rows.Count == 0)
                {
                    dtblCalcimetry = objClass.ReadExcelFile(strPath, ExcelName, "YES");
                }

                var dsStatus = (from row in dtblCalcimetry.AsEnumerable()
                                where row.Field<string>("Full Name") == WellName
                                select new
                                {
                                    Calcium = row.Field<double>("Calcium (%)"),
                                    Carbonates = row.Field<double>("Carbonates (%)")
                                });


                dtbl.Columns.Add("Calcium");
                dtbl.Columns.Add("Carbonates");
                DataRow workRow;
                foreach (var item in dsStatus)
                {
                    workRow = dtbl.NewRow();
                    workRow["Calcium"] = item.Calcium;
                    workRow["Carbonates"] = item.Carbonates;
                    dtbl.Rows.Add(workRow);
                }
                ExcelName = "Drilling Parameters";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return dtbl;
        }

        private DataTable ReadExcelDrillingParameters(string WellName)
        {
            DataTable dtbl = new DataTable(WellName);
            try
            {
                ReadFromExcel objClass = new ReadFromExcel();
                string strPath = "";
                string ExcelName = "Drilling Parameters";
                strPath = Properties.Settings.Default.DrillingParametersPath;
                if (dtblDrillingParameter.Rows.Count == 0)
                {
                    dtblDrillingParameter = objClass.ReadExcelFile(strPath, ExcelName, "YES");
                }

                var dsStatus = (from row in dtblDrillingParameter.AsEnumerable()
                                where row.Field<string>("Full Name") == WellName
                                select new
                                {
                                    WaitOnBit = row.Field<double>("Weight on Bit (k_daN)"),
                                    RPM = row.Field<double>("RPM (rpm)"),
                                    TorqAmps = row.Field<double>("Torq (amps)"),
                                    PumpPrussure = row.Field<double>("Pump Pressure (kPa)"),
                                    PumpFlow = row.Field<double>("Pump Flow (m3/min)"),
                                    MaxTorq = row.Field<double>("Max Torq (amps)"),
                                });

                dtbl.Columns.Add("Weight on Bit (k_daN)");
                dtbl.Columns.Add("RPM (rpm)");
                dtbl.Columns.Add("Torq (amps)");
                dtbl.Columns.Add("Pump Pressure (kPa)");
                dtbl.Columns.Add("Pump Flow (m3/min)");
                dtbl.Columns.Add("Max Torq (amps)");
                DataRow workRow;
                foreach (var item in dsStatus)
                {
                    workRow = dtbl.NewRow();
                    workRow["Weight on Bit (k_daN)"] = item.WaitOnBit;
                    workRow["RPM (rpm)"] = item.RPM;
                    workRow["Torq (amps)"] = item.TorqAmps;
                    workRow["Pump Pressure (kPa)"] = item.PumpPrussure;
                    workRow["Pump Flow (m3/min)"] = item.PumpFlow;
                    workRow["Max Torq (amps)"] = item.MaxTorq;
                    dtbl.Rows.Add(workRow);
                }
                ExcelName = "Drilling Parameters";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return dtbl;
        }


        private DataTable ReadIntegretedLithology(string WellName)
        {
            DataTable dtbl = new DataTable(WellName);
            try
            {
                ReadFromExcel objClass = new ReadFromExcel();
                string strPath = "";
                string ExcelName = "Integreted Lithology";
                strPath = Properties.Settings.Default.InterpretedLithologyPath;
                if (dtblIntegrtdLithology.Rows.Count == 0)
                {
                    dtblIntegrtdLithology = objClass.ReadExcelFile(strPath, ExcelName, "YES");
                }

                var dsStatus = (from row in dtblIntegrtdLithology.AsEnumerable()
                                where row.Field<string>("Full Name") == WellName
                                select new
                                {
                                    LithologyCode = row.Field<double>("Lithology Code"),
                                    Description = row.Field<double>("Description")
                                });

                dtbl.Columns.Add("Lithology Code");
                dtbl.Columns.Add("Description");
                DataRow workRow;
                foreach (var item in dsStatus)
                {
                    workRow = dtbl.NewRow();
                    workRow["Lithology Code"] = item.LithologyCode;
                    workRow["Description"] = item.Description;
                    dtbl.Rows.Add(workRow);
                }
                ExcelName = "Integreted Lithology";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return dtbl;
        }

        private DataTable ReadFromROP(string WellName)
        {
            DataTable dtbl = new DataTable(WellName);
            try
            {
                ReadFromExcel objClass = new ReadFromExcel();
                string strPath = "";
                string ExcelName = "ROP";
                strPath = Properties.Settings.Default.RopPath;
                if (dtblRop.Rows.Count == 0)
                {
                    dtblRop = objClass.ReadExcelFile(strPath, ExcelName, "YES");
                }

                var dsStatus = (from row in dtblRop.AsEnumerable()
                                where row.Field<string>("Full Name") == WellName
                                select new
                                {
                                    ROPMin = row.Field<double>("ROP (min/m)")
                                });

                dtbl.Columns.Add("ROP (min/m)");
                DataRow workRow;
                foreach (var item in dsStatus)
                {
                    workRow = dtbl.NewRow();
                    workRow["Lithology Code"] = item.ROPMin;
                    dtbl.Rows.Add(workRow);
                }
                ExcelName = "ROP";
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReadFromROP: " + ex.Message.ToString(), "EPMS -Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return dtbl;
        }

        private DataTable ReadFromHCIndicator(string WellName)
        {
            DataTable dtbl = new DataTable(WellName);
            try
            {
                ReadFromExcel objClass = new ReadFromExcel();
                string strPath = "";
                string ExcelName = "ROP";
                strPath = Properties.Settings.Default.HCIndicatorNorthPath;
                if (dtblHcIndicator.Rows.Count == 0)
                {
                    dtblHcIndicator = objClass.ReadExcelFile(strPath, ExcelName, "YES");
                }

                var dsStatus = (from row in dtblHcIndicator.AsEnumerable()
                                where row.Field<string>("Full Name") == WellName
                                select new
                                {
                                    Hcind = row.Field<double>("Hcind")
                                });

                dtbl.Columns.Add("Hcind");
                DataRow workRow;
                foreach (var item in dsStatus)
                {
                    workRow = dtbl.NewRow();
                    workRow["Hcind"] = item.Hcind;
                    dtbl.Rows.Add(workRow);
                }
                ExcelName = "HCIndicator";
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReadFromHCIndicator: " + ex.Message.ToString(), "EPMS -Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return dtbl;
        }

        private DataTable ReadFromLithologyPercentage(string WellName)
        {
            DataTable dtbl = new DataTable(WellName);
            try
            {
                ReadFromExcel objClass = new ReadFromExcel();
                string strPath = "";
                string ExcelName = "LithologyPercentage";
                strPath = Properties.Settings.Default.HCIndicatorNorthPath;
                if (dtblLithologyPercntg.Rows.Count == 0)
                {
                    dtblLithologyPercntg = objClass.ReadExcelFile(strPath, ExcelName, "YES");
                }

                var dsStatus = (from row in dtblHcIndicator.AsEnumerable()
                                where row.Field<string>("Full Name") == WellName
                                select new
                                {
                                    LithoPercntg = row.Field<double>("Lithology Percentage")
                                });

                dtbl.Columns.Add("Lithology Percentage");
                DataRow workRow;
                foreach (var item in dsStatus)
                {
                    workRow = dtbl.NewRow();
                    workRow["Lithology Percentage"] = item.LithoPercntg;
                    dtbl.Rows.Add(workRow);
                }
                ExcelName = "LithologyPercentage";
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReadFromLithologyPercentage: " + ex.Message.ToString(), "EPMS -Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return dtbl;
        }

        private DataTable MergeToFinalTable(string strTableName, DataTable dtblLas, DataTable dtblCalcimetry, DataTable dtblDrilling, DataTable dtblInterpLithlgy, DataTable dtbRop, DataTable dtblHcIndictr, DataTable dtblLithoPer)
        {
            try
            {
                if (dtblCalcimetry.Rows.Count > 0 && dtblCalcimetry.Columns.Count > 0)
                {
                    for (int i = 0; i < dtblCalcimetry.Columns.Count; i++)
                    {
                        dtblLas.Columns.Add(dtblCalcimetry.Columns[i].ToString());
                    }
                    dtblLas.Merge(dtblCalcimetry);
                }

                if (dtblDrilling.Rows.Count > 0 && dtblDrilling.Columns.Count > 0)
                {
                    for (int i = 0; i < dtblDrilling.Columns.Count; i++)
                    {
                        dtblLas.Columns.Add(dtblDrilling.Columns[i].ToString());
                    }
                    dtblLas.Merge(dtblDrilling);
                }

                if (dtblInterpLithlgy.Rows.Count > 0 && dtblInterpLithlgy.Columns.Count > 0)
                {
                    for (int i = 0; i < dtblInterpLithlgy.Columns.Count; i++)
                    {
                        dtblLas.Columns.Add(dtblInterpLithlgy.Columns[i].ToString());
                    }
                    dtblLas.Merge(dtblInterpLithlgy);
                }

                if (dtbRop.Rows.Count > 0 && dtbRop.Columns.Count > 0)
                {
                    for (int i = 0; i < dtbRop.Columns.Count; i++)
                    {
                        dtblLas.Columns.Add(dtbRop.Columns[i].ToString());
                    }
                    dtblLas.Merge(dtbRop);
                }

                if (dtblHcIndictr.Rows.Count > 0 && dtblHcIndictr.Columns.Count > 0)
                {
                    for (int i = 0; i < dtblHcIndictr.Columns.Count; i++)
                    {
                        dtblLas.Columns.Add(dtblHcIndictr.Columns[i].ToString());
                    }
                    dtblLas.Merge(dtblHcIndictr);
                }

                if (dtblLithoPer.Rows.Count > 0 && dtblLithoPer.Columns.Count > 0)
                {
                    for (int i = 0; i < dtblLithoPer.Columns.Count; i++)
                    {
                        dtblLas.Columns.Add(dtblLithoPer.Columns[i].ToString());
                    }
                    dtblLas.Merge(dtblLithoPer);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return dtblLas;
        }
    }
}
