using System;
using System.Windows.Forms;


namespace ExcelImport
{
    public partial class frmMdi : Form
    {
        public frmMdi()
        {
            InitializeComponent();
        }
        private void lasReadToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                frmLasReadandExportExcel frm = new frmLasReadandExportExcel();
                frmLasReadandExportExcel open = Application.OpenForms["frmLasReadandExportExcel"] as frmLasReadandExportExcel;
                if (open == null)
                {
                    frm.MdiParent = this;
                    frm.Show();
                }
                else
                {
                    open.Activate();
                    if (open.WindowState == FormWindowState.Minimized)
                    {
                        open.WindowState = FormWindowState.Normal;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("MDI 1: " + ex.Message, "EPMS", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }       
    }
}
