namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    partial class CountriesNowSpace_UserControl
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
#if !NETFRAMEWORK
                _ownedServices?.Dispose();
                _ownedServices = null;
#endif
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            pnlHeader            = new System.Windows.Forms.Panel();
            btnLoadCountriesNow  = new System.Windows.Forms.Button();
            lblTitle             = new System.Windows.Forms.Label();
            lblCountriesSource   = new System.Windows.Forms.Label();
            btnClearDb           = new System.Windows.Forms.Button();
            pnlExport            = new System.Windows.Forms.Panel();
            lblExportType        = new System.Windows.Forms.Label();
            cmbExportType        = new System.Windows.Forms.ComboBox();
            lblExportLibrary     = new System.Windows.Forms.Label();
            cmbExportLibrary     = new System.Windows.Forms.ComboBox();
            lblExportFilePath    = new System.Windows.Forms.Label();
            txtExportFilePath    = new System.Windows.Forms.TextBox();
            btnBrowseExport      = new System.Windows.Forms.Button();
            btnExport            = new System.Windows.Forms.Button();
            btnClearLog          = new System.Windows.Forms.Button();
            splitContainerMain   = new System.Windows.Forms.SplitContainer();
            splitContainerCountries = new System.Windows.Forms.SplitContainer();
            dgvCountriesNow      = new System.Windows.Forms.DataGridView();
            pnlStatesHeader      = new System.Windows.Forms.Panel();
            btnLoadStates        = new System.Windows.Forms.Button();
            lblSelectedCountry   = new System.Windows.Forms.Label();
            lblStatesSource      = new System.Windows.Forms.Label();
            dgvStates            = new System.Windows.Forms.DataGridView();
            rtbExportLog         = new System.Windows.Forms.RichTextBox();

            pnlHeader.SuspendLayout();
            pnlExport.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainerMain).BeginInit();
            splitContainerMain.Panel1.SuspendLayout();
            splitContainerMain.Panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainerCountries).BeginInit();
            splitContainerCountries.Panel1.SuspendLayout();
            splitContainerCountries.Panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvCountriesNow).BeginInit();
            pnlStatesHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvStates).BeginInit();
            SuspendLayout();

            // pnlHeader
            pnlHeader.Controls.Add(btnLoadCountriesNow);
            pnlHeader.Controls.Add(btnClearDb);
            pnlHeader.Controls.Add(lblCountriesSource);
            pnlHeader.Controls.Add(lblTitle);
            pnlHeader.Dock = System.Windows.Forms.DockStyle.Top;
            pnlHeader.Name = "pnlHeader";
            pnlHeader.Size = new System.Drawing.Size(1200, 40);
            pnlHeader.TabIndex = 0;

            // lblTitle
            lblTitle.AutoSize = true;
            lblTitle.Location = new System.Drawing.Point(8, 12);
            lblTitle.Name = "lblTitle";
            lblTitle.Text = "Countries Now API — countriesnow.space";

            // lblCountriesSource
            lblCountriesSource.AutoSize = true;
            lblCountriesSource.Location = new System.Drawing.Point(320, 12);
            lblCountriesSource.Name = "lblCountriesSource";
            lblCountriesSource.Text = string.Empty;

            // btnClearDb
            btnClearDb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            btnClearDb.Location = new System.Drawing.Point(1082, 8);
            btnClearDb.Name = "btnClearDb";
            btnClearDb.Size = new System.Drawing.Size(108, 26);
            btnClearDb.TabIndex = 1;
            btnClearDb.Text = "Clear Database";
            btnClearDb.UseVisualStyleBackColor = true;
            btnClearDb.Click += new System.EventHandler(btnClearDb_Click);

            // btnLoadCountriesNow
            btnLoadCountriesNow.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            btnLoadCountriesNow.Location = new System.Drawing.Point(968, 8);
            btnLoadCountriesNow.Name = "btnLoadCountriesNow";
            btnLoadCountriesNow.Size = new System.Drawing.Size(108, 26);
            btnLoadCountriesNow.TabIndex = 0;
            btnLoadCountriesNow.Text = "Load Countries";
            btnLoadCountriesNow.UseVisualStyleBackColor = true;
            btnLoadCountriesNow.Click += new System.EventHandler(btnLoadCountriesNow_Click);

            // pnlExport
            pnlExport.Controls.Add(lblExportType);
            pnlExport.Controls.Add(cmbExportType);
            pnlExport.Controls.Add(lblExportLibrary);
            pnlExport.Controls.Add(cmbExportLibrary);
            pnlExport.Controls.Add(lblExportFilePath);
            pnlExport.Controls.Add(txtExportFilePath);
            pnlExport.Controls.Add(btnBrowseExport);
            pnlExport.Controls.Add(btnExport);
            pnlExport.Controls.Add(btnClearLog);
            pnlExport.Dock = System.Windows.Forms.DockStyle.Bottom;
            pnlExport.Name = "pnlExport";
            pnlExport.Size = new System.Drawing.Size(1200, 44);
            pnlExport.TabIndex = 1;

            // lblExportType
            lblExportType.AutoSize = true;
            lblExportType.Location = new System.Drawing.Point(8, 14);
            lblExportType.Name = "lblExportType";
            lblExportType.Text = "Export Type:";

            // cmbExportType
            cmbExportType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbExportType.Location = new System.Drawing.Point(86, 10);
            cmbExportType.Name = "cmbExportType";
            cmbExportType.Size = new System.Drawing.Size(120, 23);
            cmbExportType.TabIndex = 0;

            // lblExportLibrary
            lblExportLibrary.AutoSize = true;
            lblExportLibrary.Location = new System.Drawing.Point(214, 14);
            lblExportLibrary.Name = "lblExportLibrary";
            lblExportLibrary.Text = "Library:";

            // cmbExportLibrary
            cmbExportLibrary.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbExportLibrary.Location = new System.Drawing.Point(262, 10);
            cmbExportLibrary.Name = "cmbExportLibrary";
            cmbExportLibrary.Size = new System.Drawing.Size(100, 23);
            cmbExportLibrary.TabIndex = 1;

            // lblExportFilePath
            lblExportFilePath.AutoSize = true;
            lblExportFilePath.Location = new System.Drawing.Point(370, 14);
            lblExportFilePath.Name = "lblExportFilePath";
            lblExportFilePath.Text = "File:";

            // txtExportFilePath
            txtExportFilePath.Anchor = System.Windows.Forms.AnchorStyles.Top
                | System.Windows.Forms.AnchorStyles.Left
                | System.Windows.Forms.AnchorStyles.Right;
            txtExportFilePath.Location = new System.Drawing.Point(400, 10);
            txtExportFilePath.Name = "txtExportFilePath";
            txtExportFilePath.Size = new System.Drawing.Size(498, 23);
            txtExportFilePath.TabIndex = 2;

            // btnBrowseExport
            btnBrowseExport.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            btnBrowseExport.Location = new System.Drawing.Point(906, 8);
            btnBrowseExport.Name = "btnBrowseExport";
            btnBrowseExport.Size = new System.Drawing.Size(90, 26);
            btnBrowseExport.TabIndex = 3;
            btnBrowseExport.Text = "Browse...";
            btnBrowseExport.UseVisualStyleBackColor = true;
            btnBrowseExport.Click += new System.EventHandler(btnBrowseExport_Click);

            // btnExport
            btnExport.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            btnExport.Location = new System.Drawing.Point(1004, 8);
            btnExport.Name = "btnExport";
            btnExport.Size = new System.Drawing.Size(100, 26);
            btnExport.TabIndex = 4;
            btnExport.Text = "Export";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += new System.EventHandler(btnExport_Click);

            // btnClearLog
            btnClearLog.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            btnClearLog.Location = new System.Drawing.Point(1112, 8);
            btnClearLog.Name = "btnClearLog";
            btnClearLog.Size = new System.Drawing.Size(80, 26);
            btnClearLog.TabIndex = 5;
            btnClearLog.Text = "Clear Log";
            btnClearLog.UseVisualStyleBackColor = true;
            btnClearLog.Click += new System.EventHandler(btnClearLog_Click);

            // splitContainerMain — top panel: countries/states; bottom panel: export log
            splitContainerMain.Dock = System.Windows.Forms.DockStyle.Fill;
            splitContainerMain.Name = "splitContainerMain";
            splitContainerMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            splitContainerMain.SplitterDistance = 540;
            splitContainerMain.TabIndex = 2;

            // splitContainerMain.Panel1 — countries + states split
            splitContainerCountries.Dock = System.Windows.Forms.DockStyle.Fill;
            splitContainerCountries.Name = "splitContainerCountries";
            splitContainerCountries.SplitterDistance = 600;
            splitContainerCountries.TabIndex = 0;

            // dgvCountriesNow
            dgvCountriesNow.AllowUserToAddRows = false;
            dgvCountriesNow.AllowUserToDeleteRows = false;
            dgvCountriesNow.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dgvCountriesNow.Dock = System.Windows.Forms.DockStyle.Fill;
            dgvCountriesNow.MultiSelect = false;
            dgvCountriesNow.Name = "dgvCountriesNow";
            dgvCountriesNow.ReadOnly = true;
            dgvCountriesNow.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            dgvCountriesNow.TabIndex = 0;
            dgvCountriesNow.SelectionChanged += new System.EventHandler(dgvCountriesNow_SelectionChanged);
            splitContainerCountries.Panel1.Controls.Add(dgvCountriesNow);

            // pnlStatesHeader
            pnlStatesHeader.Controls.Add(btnLoadStates);
            pnlStatesHeader.Controls.Add(lblStatesSource);
            pnlStatesHeader.Controls.Add(lblSelectedCountry);
            pnlStatesHeader.Dock = System.Windows.Forms.DockStyle.Top;
            pnlStatesHeader.Name = "pnlStatesHeader";
            pnlStatesHeader.Size = new System.Drawing.Size(596, 40);
            pnlStatesHeader.TabIndex = 0;

            // lblSelectedCountry
            lblSelectedCountry.AutoSize = true;
            lblSelectedCountry.Location = new System.Drawing.Point(8, 12);
            lblSelectedCountry.Name = "lblSelectedCountry";
            lblSelectedCountry.Text = "Select a country to load states";

            // lblStatesSource
            lblStatesSource.AutoSize = true;
            lblStatesSource.Location = new System.Drawing.Point(250, 12);
            lblStatesSource.Name = "lblStatesSource";
            lblStatesSource.Text = string.Empty;

            // btnLoadStates
            btnLoadStates.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            btnLoadStates.Enabled = false;
            btnLoadStates.Location = new System.Drawing.Point(500, 8);
            btnLoadStates.Name = "btnLoadStates";
            btnLoadStates.Size = new System.Drawing.Size(88, 26);
            btnLoadStates.TabIndex = 0;
            btnLoadStates.Text = "Load States";
            btnLoadStates.UseVisualStyleBackColor = true;
            btnLoadStates.Click += new System.EventHandler(btnLoadStates_Click);

            // dgvStates
            dgvStates.AllowUserToAddRows = false;
            dgvStates.AllowUserToDeleteRows = false;
            dgvStates.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dgvStates.Dock = System.Windows.Forms.DockStyle.Fill;
            dgvStates.Name = "dgvStates";
            dgvStates.ReadOnly = true;
            dgvStates.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            dgvStates.TabIndex = 1;
            splitContainerCountries.Panel2.Controls.Add(dgvStates);
            splitContainerCountries.Panel2.Controls.Add(pnlStatesHeader);

            splitContainerMain.Panel1.Controls.Add(splitContainerCountries);

            // splitContainerMain.Panel2 — export log
            rtbExportLog.Dock = System.Windows.Forms.DockStyle.Fill;
            rtbExportLog.Name = "rtbExportLog";
            rtbExportLog.ReadOnly = true;
            rtbExportLog.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Both;
            rtbExportLog.TabIndex = 0;
            rtbExportLog.Font = new System.Drawing.Font("Consolas", 9F);
            rtbExportLog.BackColor = System.Drawing.Color.Black;
            rtbExportLog.ForeColor = System.Drawing.Color.Lime;
            splitContainerMain.Panel2.Controls.Add(rtbExportLog);

            // CountriesNowSpace_UserControl
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            Controls.Add(splitContainerMain);
            Controls.Add(pnlExport);
            Controls.Add(pnlHeader);
            Name = "CountriesNowSpace_UserControl";
            Size = new System.Drawing.Size(1200, 800);

            pnlHeader.ResumeLayout(false);
            pnlHeader.PerformLayout();
            pnlExport.ResumeLayout(false);
            pnlExport.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainerMain).EndInit();
            splitContainerMain.Panel1.ResumeLayout(false);
            splitContainerMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainerCountries).EndInit();
            splitContainerCountries.Panel1.ResumeLayout(false);
            splitContainerCountries.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvCountriesNow).EndInit();
            pnlStatesHeader.ResumeLayout(false);
            pnlStatesHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvStates).EndInit();
            ResumeLayout(false);
        }

        private System.Windows.Forms.Panel pnlHeader;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblCountriesSource;
        private System.Windows.Forms.Button btnLoadCountriesNow;
        private System.Windows.Forms.Button btnClearDb;
        private System.Windows.Forms.Panel pnlExport;
        private System.Windows.Forms.Label lblExportType;
        private System.Windows.Forms.ComboBox cmbExportType;
        private System.Windows.Forms.Label lblExportLibrary;
        private System.Windows.Forms.ComboBox cmbExportLibrary;
        private System.Windows.Forms.Label lblExportFilePath;
        private System.Windows.Forms.TextBox txtExportFilePath;
        private System.Windows.Forms.Button btnBrowseExport;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnClearLog;
        private System.Windows.Forms.SplitContainer splitContainerMain;
        private System.Windows.Forms.SplitContainer splitContainerCountries;
        private System.Windows.Forms.DataGridView dgvCountriesNow;
        private System.Windows.Forms.Panel pnlStatesHeader;
        private System.Windows.Forms.Button btnLoadStates;
        private System.Windows.Forms.Label lblSelectedCountry;
        private System.Windows.Forms.Label lblStatesSource;
        private System.Windows.Forms.DataGridView dgvStates;
        private System.Windows.Forms.RichTextBox rtbExportLog;
    }
}
