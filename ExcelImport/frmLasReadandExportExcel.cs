using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Diagnostics;

namespace ExcelImport
{
    public partial class frmLasReadandExportExcel : Form
    {
        frmLoading frmLodingObj = new frmLoading();
        public frmLasReadandExportExcel()
        {
            InitializeComponent();

            ToolTip toolTip1 = new ToolTip();
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(btnBrowseFolder, "Click to browse folder");
            toolTip1.SetToolTip(btnProcess, "Click this to start the process.!");

            if (Properties.Settings.Default.InputPath != string.Empty)
            {
                txtFilePath.Text = Properties.Settings.Default.InputPath;
                txtOutPutPath.Text = Properties.Settings.Default.OutputPath;
                chkRemember.Checked = Properties.Settings.Default.IspathChecked;

                txtCalcimetry.Text = Properties.Settings.Default.CalcimetryPath;
                txtDrillingParameter.Text = Properties.Settings.Default.DrillingParametersPath;
                txtInterpretedLithology.Text = Properties.Settings.Default.InterpretedLithologyPath;
                txtRop.Text = Properties.Settings.Default.RopPath;
                txtHcIndicator.Text = Properties.Settings.Default.HCIndicatorNorthPath;
                txtLithologyPercentage.Text = Properties.Settings.Default.LithologyPercntPath;
            }
        }
        private bool StartExport()
        {
            bool isResult = false;
            try
            {
                if (chkRemember.Checked)
                {
                    Properties.Settings.Default.InputPath = txtFilePath.Text;
                    Properties.Settings.Default.OutputPath = txtOutPutPath.Text;
                    Properties.Settings.Default.IspathChecked = chkRemember.Checked;


                    Properties.Settings.Default.CalcimetryPath = txtCalcimetry.Text.Trim();
                    Properties.Settings.Default.DrillingParametersPath = txtDrillingParameter.Text.Trim();
                    Properties.Settings.Default.InterpretedLithologyPath = txtInterpretedLithology.Text.Trim();
                    Properties.Settings.Default.RopPath = txtRop.Text.Trim();
                    Properties.Settings.Default.HCIndicatorNorthPath = txtHcIndicator.Text.Trim();
                    Properties.Settings.Default.LithologyPercntPath = txtLithologyPercentage.Text.Trim();



                    Properties.Settings.Default.Save();
                }
                string strBase = AppDomain.CurrentDomain.BaseDirectory;
                string strFullPath = strBase + ConfigurationManager.AppSettings["TemplatePath"];
                string strTemplateName = ConfigurationManager.AppSettings["TemplateName"];
                ReadLasAndExportDataSet objClass = new ReadLasAndExportDataSet();
                isResult = objClass.ReadFileFromFolderandProcess(txtFilePath.Text.Trim(), txtOutPutPath.Text.Trim(), strFullPath, strTemplateName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
            return isResult;
        }

        private void btnBrowseFolder_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog obj = new FolderBrowserDialog();
                DialogResult result = obj.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtFilePath.Text = obj.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }


        private void btnProcess_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtFilePath.Text != string.Empty)
                {
                    if (txtOutPutPath.Text != string.Empty)
                    {
                        bwrk1.RunWorkerAsync();
                        frmLodingObj.ShowFromOtherForms();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }

        private void bwrk1_DoWork(object sender, DoWorkEventArgs e)
        {
            //SampleExport();
            try
            {
                bool isSuccess = StartExport();
                if (isSuccess)
                {
                    DialogResult dr = MessageBox.Show("Processed Successfully. Do you want to open the output folder.?", "EPMS-Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (dr == DialogResult.OK)
                    {
                        string strOutputPath = txtOutPutPath.Text;
                        Process.Start(@strOutputPath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
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
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }

        private void frmLasReadandExportExcel_Load(object sender, EventArgs e)
        {
            //if (txtCalcimetry.Text != string.Empty && txtDrillingParameter.Text != string.Empty && txtInterpretedLithology.Text != string.Empty)
            //{
            //    btnProcess_Click(sender, e);
            //}
        }

        private void btnCalcimetry_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog obj = new OpenFileDialog();
                DialogResult result = obj.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtCalcimetry.Text = obj.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }

        private void btnDrillingParameters_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog obj = new OpenFileDialog();
                DialogResult result = obj.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtDrillingParameter.Text = obj.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }

        private void btnInterpretedLithology_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog obj = new OpenFileDialog();
                DialogResult result = obj.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtInterpretedLithology.Text = obj.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }

        private void btnRop_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog obj = new OpenFileDialog();
                DialogResult result = obj.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtRop.Text = obj.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }

        private void btnHcNorthFields_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog obj = new OpenFileDialog();
                DialogResult result = obj.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtHcIndicator.Text = obj.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }

        private void btnLithology_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog obj = new OpenFileDialog();
                DialogResult result = obj.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtLithologyPercentage.Text = obj.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "EPMS-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Common.WriteToFile(ex.Message, false);
            }
        }
    }
}
