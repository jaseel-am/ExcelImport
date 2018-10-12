namespace ExcelImport
{
    partial class frmLasReadandExportExcel
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmLasReadandExportExcel));
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtOutPutPath = new System.Windows.Forms.TextBox();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnProcess = new System.Windows.Forms.Button();
            this.bwrk1 = new System.ComponentModel.BackgroundWorker();
            this.chkRemember = new System.Windows.Forms.CheckBox();
            this.txtCalcimetry = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnBrowseFolder = new System.Windows.Forms.Button();
            this.txtDrillingParameter = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtInterpretedLithology = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtRop = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtHcIndicator = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtLithologyPercentage = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnCalcimetry = new System.Windows.Forms.Button();
            this.btnDrillingParameters = new System.Windows.Forms.Button();
            this.btnInterpretedLithology = new System.Windows.Forms.Button();
            this.btnRop = new System.Windows.Forms.Button();
            this.btnHcNorthFields = new System.Windows.Forms.Button();
            this.btnLithology = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(7, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(110, 14);
            this.label3.TabIndex = 22;
            this.label3.Text = "Las Folder Path:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(7, 223);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(125, 14);
            this.label1.TabIndex = 24;
            this.label1.Text = "Excel Output Path:";
            // 
            // txtOutPutPath
            // 
            this.txtOutPutPath.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOutPutPath.Location = new System.Drawing.Point(245, 220);
            this.txtOutPutPath.Name = "txtOutPutPath";
            this.txtOutPutPath.Size = new System.Drawing.Size(636, 22);
            this.txtOutPutPath.TabIndex = 20;
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilePath.Location = new System.Drawing.Point(245, 22);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(636, 22);
            this.txtFilePath.TabIndex = 19;
            // 
            // btnProcess
            // 
            this.btnProcess.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btnProcess.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnProcess.Font = new System.Drawing.Font("Verdana", 9F);
            this.btnProcess.ForeColor = System.Drawing.Color.White;
            this.btnProcess.Location = new System.Drawing.Point(245, 273);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(75, 27);
            this.btnProcess.TabIndex = 18;
            this.btnProcess.Text = "Process";
            this.btnProcess.UseVisualStyleBackColor = false;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // bwrk1
            // 
            this.bwrk1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bwrk1_DoWork);
            this.bwrk1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bwrk1_RunWorkerCompleted);
            // 
            // chkRemember
            // 
            this.chkRemember.AutoSize = true;
            this.chkRemember.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkRemember.ForeColor = System.Drawing.Color.White;
            this.chkRemember.Location = new System.Drawing.Point(245, 249);
            this.chkRemember.Name = "chkRemember";
            this.chkRemember.Size = new System.Drawing.Size(110, 18);
            this.chkRemember.TabIndex = 27;
            this.chkRemember.Text = "Remember Path";
            this.chkRemember.UseVisualStyleBackColor = true;
            // 
            // txtCalcimetry
            // 
            this.txtCalcimetry.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCalcimetry.Location = new System.Drawing.Point(245, 50);
            this.txtCalcimetry.Name = "txtCalcimetry";
            this.txtCalcimetry.Size = new System.Drawing.Size(636, 22);
            this.txtCalcimetry.TabIndex = 19;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(7, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 14);
            this.label2.TabIndex = 22;
            this.label2.Text = "Calcimetry:";
            // 
            // btnBrowseFolder
            // 
            this.btnBrowseFolder.BackColor = System.Drawing.Color.SteelBlue;
            this.btnBrowseFolder.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnBrowseFolder.BackgroundImage")));
            this.btnBrowseFolder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowseFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowseFolder.ForeColor = System.Drawing.Color.SteelBlue;
            this.btnBrowseFolder.Location = new System.Drawing.Point(882, 18);
            this.btnBrowseFolder.Name = "btnBrowseFolder";
            this.btnBrowseFolder.Size = new System.Drawing.Size(37, 27);
            this.btnBrowseFolder.TabIndex = 28;
            this.btnBrowseFolder.UseVisualStyleBackColor = false;
            this.btnBrowseFolder.Click += new System.EventHandler(this.btnBrowseFolder_Click);
            // 
            // txtDrillingParameter
            // 
            this.txtDrillingParameter.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDrillingParameter.Location = new System.Drawing.Point(245, 78);
            this.txtDrillingParameter.Name = "txtDrillingParameter";
            this.txtDrillingParameter.Size = new System.Drawing.Size(636, 22);
            this.txtDrillingParameter.TabIndex = 19;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(7, 83);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(131, 14);
            this.label4.TabIndex = 22;
            this.label4.Text = "Drilling Parameters:";
            // 
            // txtInterpretedLithology
            // 
            this.txtInterpretedLithology.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtInterpretedLithology.Location = new System.Drawing.Point(245, 106);
            this.txtInterpretedLithology.Name = "txtInterpretedLithology";
            this.txtInterpretedLithology.Size = new System.Drawing.Size(636, 22);
            this.txtInterpretedLithology.TabIndex = 19;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Transparent;
            this.label5.Location = new System.Drawing.Point(7, 111);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(146, 14);
            this.label5.TabIndex = 22;
            this.label5.Text = "Interpreted Lithology:";
            // 
            // txtRop
            // 
            this.txtRop.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRop.Location = new System.Drawing.Point(245, 134);
            this.txtRop.Name = "txtRop";
            this.txtRop.Size = new System.Drawing.Size(636, 22);
            this.txtRop.TabIndex = 19;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Transparent;
            this.label6.Location = new System.Drawing.Point(7, 140);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(38, 14);
            this.label6.TabIndex = 22;
            this.label6.Text = "ROP:";
            // 
            // txtHcIndicator
            // 
            this.txtHcIndicator.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHcIndicator.Location = new System.Drawing.Point(245, 162);
            this.txtHcIndicator.Name = "txtHcIndicator";
            this.txtHcIndicator.Size = new System.Drawing.Size(636, 22);
            this.txtHcIndicator.TabIndex = 19;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Transparent;
            this.label7.Location = new System.Drawing.Point(7, 167);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(174, 14);
            this.label7.TabIndex = 22;
            this.label7.Text = "HC Indicator NORTH fields:";
            // 
            // txtLithologyPercentage
            // 
            this.txtLithologyPercentage.Font = new System.Drawing.Font("Cambria", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLithologyPercentage.Location = new System.Drawing.Point(245, 190);
            this.txtLithologyPercentage.Name = "txtLithologyPercentage";
            this.txtLithologyPercentage.Size = new System.Drawing.Size(636, 22);
            this.txtLithologyPercentage.TabIndex = 19;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.Transparent;
            this.label8.Location = new System.Drawing.Point(7, 195);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(229, 14);
            this.label8.TabIndex = 22;
            this.label8.Text = "Lithology percentage NORTH fields:";
            // 
            // btnCalcimetry
            // 
            this.btnCalcimetry.BackColor = System.Drawing.Color.SteelBlue;
            this.btnCalcimetry.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnCalcimetry.BackgroundImage")));
            this.btnCalcimetry.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnCalcimetry.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCalcimetry.ForeColor = System.Drawing.Color.SteelBlue;
            this.btnCalcimetry.Location = new System.Drawing.Point(882, 46);
            this.btnCalcimetry.Name = "btnCalcimetry";
            this.btnCalcimetry.Size = new System.Drawing.Size(37, 27);
            this.btnCalcimetry.TabIndex = 28;
            this.btnCalcimetry.UseVisualStyleBackColor = false;
            this.btnCalcimetry.Click += new System.EventHandler(this.btnCalcimetry_Click);
            // 
            // btnDrillingParameters
            // 
            this.btnDrillingParameters.BackColor = System.Drawing.Color.SteelBlue;
            this.btnDrillingParameters.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnDrillingParameters.BackgroundImage")));
            this.btnDrillingParameters.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnDrillingParameters.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDrillingParameters.ForeColor = System.Drawing.Color.SteelBlue;
            this.btnDrillingParameters.Location = new System.Drawing.Point(882, 74);
            this.btnDrillingParameters.Name = "btnDrillingParameters";
            this.btnDrillingParameters.Size = new System.Drawing.Size(37, 27);
            this.btnDrillingParameters.TabIndex = 28;
            this.btnDrillingParameters.UseVisualStyleBackColor = false;
            this.btnDrillingParameters.Click += new System.EventHandler(this.btnDrillingParameters_Click);
            // 
            // btnInterpretedLithology
            // 
            this.btnInterpretedLithology.BackColor = System.Drawing.Color.SteelBlue;
            this.btnInterpretedLithology.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnInterpretedLithology.BackgroundImage")));
            this.btnInterpretedLithology.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnInterpretedLithology.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInterpretedLithology.ForeColor = System.Drawing.Color.SteelBlue;
            this.btnInterpretedLithology.Location = new System.Drawing.Point(882, 102);
            this.btnInterpretedLithology.Name = "btnInterpretedLithology";
            this.btnInterpretedLithology.Size = new System.Drawing.Size(37, 27);
            this.btnInterpretedLithology.TabIndex = 28;
            this.btnInterpretedLithology.UseVisualStyleBackColor = false;
            this.btnInterpretedLithology.Click += new System.EventHandler(this.btnInterpretedLithology_Click);
            // 
            // btnRop
            // 
            this.btnRop.BackColor = System.Drawing.Color.SteelBlue;
            this.btnRop.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnRop.BackgroundImage")));
            this.btnRop.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRop.ForeColor = System.Drawing.Color.SteelBlue;
            this.btnRop.Location = new System.Drawing.Point(882, 134);
            this.btnRop.Name = "btnRop";
            this.btnRop.Size = new System.Drawing.Size(37, 27);
            this.btnRop.TabIndex = 28;
            this.btnRop.UseVisualStyleBackColor = false;
            this.btnRop.Click += new System.EventHandler(this.btnRop_Click);
            // 
            // btnHcNorthFields
            // 
            this.btnHcNorthFields.BackColor = System.Drawing.Color.SteelBlue;
            this.btnHcNorthFields.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnHcNorthFields.BackgroundImage")));
            this.btnHcNorthFields.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnHcNorthFields.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnHcNorthFields.ForeColor = System.Drawing.Color.SteelBlue;
            this.btnHcNorthFields.Location = new System.Drawing.Point(882, 162);
            this.btnHcNorthFields.Name = "btnHcNorthFields";
            this.btnHcNorthFields.Size = new System.Drawing.Size(37, 27);
            this.btnHcNorthFields.TabIndex = 28;
            this.btnHcNorthFields.UseVisualStyleBackColor = false;
            this.btnHcNorthFields.Click += new System.EventHandler(this.btnHcNorthFields_Click);
            // 
            // btnLithology
            // 
            this.btnLithology.BackColor = System.Drawing.Color.SteelBlue;
            this.btnLithology.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnLithology.BackgroundImage")));
            this.btnLithology.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnLithology.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLithology.ForeColor = System.Drawing.Color.SteelBlue;
            this.btnLithology.Location = new System.Drawing.Point(882, 189);
            this.btnLithology.Name = "btnLithology";
            this.btnLithology.Size = new System.Drawing.Size(37, 27);
            this.btnLithology.TabIndex = 28;
            this.btnLithology.UseVisualStyleBackColor = false;
            this.btnLithology.Click += new System.EventHandler(this.btnLithology_Click);
            // 
            // frmLasReadandExportExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SteelBlue;
            this.ClientSize = new System.Drawing.Size(933, 377);
            this.Controls.Add(this.btnLithology);
            this.Controls.Add(this.btnHcNorthFields);
            this.Controls.Add(this.btnRop);
            this.Controls.Add(this.btnInterpretedLithology);
            this.Controls.Add(this.btnDrillingParameters);
            this.Controls.Add(this.btnCalcimetry);
            this.Controls.Add(this.btnBrowseFolder);
            this.Controls.Add(this.chkRemember);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtOutPutPath);
            this.Controls.Add(this.txtLithologyPercentage);
            this.Controls.Add(this.txtHcIndicator);
            this.Controls.Add(this.txtRop);
            this.Controls.Add(this.txtInterpretedLithology);
            this.Controls.Add(this.txtDrillingParameter);
            this.Controls.Add(this.txtCalcimetry);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.btnProcess);
            this.MaximizeBox = false;
            this.Name = "frmLasReadandExportExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = ".Las Read and Export Excel";
            this.Load += new System.EventHandler(this.frmLasReadandExportExcel_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtOutPutPath;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnProcess;
        private System.ComponentModel.BackgroundWorker bwrk1;
        private System.Windows.Forms.CheckBox chkRemember;
        private System.Windows.Forms.TextBox txtCalcimetry;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnBrowseFolder;
        private System.Windows.Forms.TextBox txtDrillingParameter;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtInterpretedLithology;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtRop;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtHcIndicator;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtLithologyPercentage;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnCalcimetry;
        private System.Windows.Forms.Button btnDrillingParameters;
        private System.Windows.Forms.Button btnInterpretedLithology;
        private System.Windows.Forms.Button btnRop;
        private System.Windows.Forms.Button btnHcNorthFields;
        private System.Windows.Forms.Button btnLithology;
    }
}