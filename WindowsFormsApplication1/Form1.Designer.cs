namespace ExcelTranscriptionMachine
{
    partial class Form1
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
            copier.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnGo = new System.Windows.Forms.Button();
            this.textBoxClaimNum = new System.Windows.Forms.TextBox();
            this.lblClaimNum = new System.Windows.Forms.Label();
            this.progressBarClaims = new System.Windows.Forms.ProgressBar();
            this.openFileDialogBilling = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialogResponse = new System.Windows.Forms.OpenFileDialog();
            this.buttonPopBillingFileDialog = new System.Windows.Forms.Button();
            this.buttonPopResponseFileDialog = new System.Windows.Forms.Button();
            this.textBoxBillingFileName = new AutoEllipsis.TextBoxEllipsis();
            this.textBoxResponseFileName = new AutoEllipsis.TextBoxEllipsis();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lblBillingFile = new System.Windows.Forms.Label();
            this.lblResponseFile = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGo
            // 
            this.btnGo.Location = new System.Drawing.Point(390, 99);
            this.btnGo.Name = "btnGo";
            this.btnGo.Size = new System.Drawing.Size(75, 23);
            this.btnGo.TabIndex = 0;
            this.btnGo.Text = "Begin";
            this.btnGo.UseVisualStyleBackColor = true;
            this.btnGo.Click += new System.EventHandler(this.btnGo_Click);
            // 
            // textBoxClaimNum
            // 
            this.textBoxClaimNum.Enabled = false;
            this.textBoxClaimNum.Location = new System.Drawing.Point(82, 161);
            this.textBoxClaimNum.Name = "textBoxClaimNum";
            this.textBoxClaimNum.Size = new System.Drawing.Size(172, 20);
            this.textBoxClaimNum.TabIndex = 1;
            // 
            // lblClaimNum
            // 
            this.lblClaimNum.AutoSize = true;
            this.lblClaimNum.Location = new System.Drawing.Point(34, 164);
            this.lblClaimNum.Name = "lblClaimNum";
            this.lblClaimNum.Size = new System.Drawing.Size(42, 13);
            this.lblClaimNum.TabIndex = 2;
            this.lblClaimNum.Text = "Claim #";
            // 
            // progressBarClaims
            // 
            this.progressBarClaims.Location = new System.Drawing.Point(37, 197);
            this.progressBarClaims.Name = "progressBarClaims";
            this.progressBarClaims.Size = new System.Drawing.Size(428, 23);
            this.progressBarClaims.TabIndex = 3;
            // 
            // openFileDialogBilling
            // 
            this.openFileDialogBilling.Filter = "Excel 1997-2003|*.xls|All files|*.*";
            // 
            // openFileDialogResponse
            // 
            this.openFileDialogResponse.Filter = "Excel 1997-2003|*.xls|All files|*.*";
            // 
            // buttonPopBillingFileDialog
            // 
            this.buttonPopBillingFileDialog.Location = new System.Drawing.Point(434, 41);
            this.buttonPopBillingFileDialog.Name = "buttonPopBillingFileDialog";
            this.buttonPopBillingFileDialog.Size = new System.Drawing.Size(31, 23);
            this.buttonPopBillingFileDialog.TabIndex = 4;
            this.buttonPopBillingFileDialog.Text = "...";
            this.buttonPopBillingFileDialog.UseVisualStyleBackColor = true;
            this.buttonPopBillingFileDialog.Click += new System.EventHandler(this.buttonPopBillingFileDialog_Click);
            // 
            // buttonPopResponseFileDialog
            // 
            this.buttonPopResponseFileDialog.Location = new System.Drawing.Point(434, 71);
            this.buttonPopResponseFileDialog.Name = "buttonPopResponseFileDialog";
            this.buttonPopResponseFileDialog.Size = new System.Drawing.Size(31, 23);
            this.buttonPopResponseFileDialog.TabIndex = 5;
            this.buttonPopResponseFileDialog.Text = "...";
            this.buttonPopResponseFileDialog.UseVisualStyleBackColor = true;
            this.buttonPopResponseFileDialog.Click += new System.EventHandler(this.buttonPopResponseFileDialog_Click);
            // 
            // textBoxBillingFileName
            // 
            this.textBoxBillingFileName.AutoEllipsis = ((AutoEllipsis.EllipsisFormat)((AutoEllipsis.EllipsisFormat.Start | AutoEllipsis.EllipsisFormat.Path)));
            this.textBoxBillingFileName.Location = new System.Drawing.Point(138, 44);
            this.textBoxBillingFileName.Name = "textBoxBillingFileName";
            this.textBoxBillingFileName.Size = new System.Drawing.Size(289, 20);
            this.textBoxBillingFileName.TabIndex = 6;
            this.textBoxBillingFileName.MouseClick += new System.Windows.Forms.MouseEventHandler(this.textBoxBillingFileName_MouseClick);
            this.textBoxBillingFileName.TextChanged += new System.EventHandler(this.textBoxBillingFileName_TextChanged);
            // 
            // textBoxResponseFileName
            // 
            this.textBoxResponseFileName.AutoEllipsis = ((AutoEllipsis.EllipsisFormat)((AutoEllipsis.EllipsisFormat.Start | AutoEllipsis.EllipsisFormat.Path)));
            this.textBoxResponseFileName.Location = new System.Drawing.Point(138, 73);
            this.textBoxResponseFileName.Name = "textBoxResponseFileName";
            this.textBoxResponseFileName.Size = new System.Drawing.Size(289, 20);
            this.textBoxResponseFileName.TabIndex = 7;
            this.textBoxResponseFileName.MouseClick += new System.Windows.Forms.MouseEventHandler(this.textBoxResponseFileName_MouseClick);
            this.textBoxResponseFileName.TextChanged += new System.EventHandler(this.textBoxResponseFileName_TextChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(500, 24);
            this.menuStrip1.TabIndex = 8;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // lblBillingFile
            // 
            this.lblBillingFile.AutoSize = true;
            this.lblBillingFile.Location = new System.Drawing.Point(37, 51);
            this.lblBillingFile.Name = "lblBillingFile";
            this.lblBillingFile.Size = new System.Drawing.Size(50, 13);
            this.lblBillingFile.TabIndex = 9;
            this.lblBillingFile.Text = "Billing file";
            // 
            // lblResponseFile
            // 
            this.lblResponseFile.AutoSize = true;
            this.lblResponseFile.Location = new System.Drawing.Point(37, 76);
            this.lblResponseFile.Name = "lblResponseFile";
            this.lblResponseFile.Size = new System.Drawing.Size(71, 13);
            this.lblResponseFile.TabIndex = 10;
            this.lblResponseFile.Text = "Response file";
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(390, 157);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 312);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblResponseFile);
            this.Controls.Add(this.lblBillingFile);
            this.Controls.Add(this.textBoxResponseFileName);
            this.Controls.Add(this.textBoxBillingFileName);
            this.Controls.Add(this.buttonPopResponseFileDialog);
            this.Controls.Add(this.buttonPopBillingFileDialog);
            this.Controls.Add(this.progressBarClaims);
            this.Controls.Add(this.lblClaimNum);
            this.Controls.Add(this.textBoxClaimNum);
            this.Controls.Add(this.btnGo);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Excel Transcription Machine";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGo;
        private System.Windows.Forms.TextBox textBoxClaimNum;
        private System.Windows.Forms.Label lblClaimNum;
        private System.Windows.Forms.ProgressBar progressBarClaims;
        private System.Windows.Forms.OpenFileDialog openFileDialogBilling;
        private System.Windows.Forms.OpenFileDialog openFileDialogResponse;
        private System.Windows.Forms.Button buttonPopBillingFileDialog;
        private System.Windows.Forms.Button buttonPopResponseFileDialog;
        private AutoEllipsis.TextBoxEllipsis textBoxBillingFileName;
        private AutoEllipsis.TextBoxEllipsis textBoxResponseFileName;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.Label lblBillingFile;
        private System.Windows.Forms.Label lblResponseFile;
        private System.Windows.Forms.Button btnCancel;
    }
}

