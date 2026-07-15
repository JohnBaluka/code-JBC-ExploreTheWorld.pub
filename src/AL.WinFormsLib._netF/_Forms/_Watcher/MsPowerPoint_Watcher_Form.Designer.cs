namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    partial class MsPowerPoint_Watcher_Form
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            pnlHeader            = new System.Windows.Forms.Panel();
            btnConnectDisconnect = new System.Windows.Forms.Button();
            splitContainer       = new System.Windows.Forms.SplitContainer();
            pnlLogTop            = new System.Windows.Forms.Panel();
            btnClearLog          = new System.Windows.Forms.Button();
            rtbLog               = new System.Windows.Forms.RichTextBox();
            dgvEvents            = new System.Windows.Forms.DataGridView();
            colName              = new System.Windows.Forms.DataGridViewTextBoxColumn();
            colCategory          = new System.Windows.Forms.DataGridViewTextBoxColumn();
            colLog               = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            pnlFooter            = new System.Windows.Forms.Panel();
            btnSelectOutputFile  = new System.Windows.Forms.Button();
            lblJsonWriteMethod   = new System.Windows.Forms.Label();
            cbxJsonWriteMethod   = new System.Windows.Forms.ComboBox();
            txtOutputFilePath    = new System.Windows.Forms.TextBox();
            btnSaveAsJson        = new System.Windows.Forms.Button();

            pnlHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer).BeginInit();
            splitContainer.Panel1.SuspendLayout();
            splitContainer.Panel2.SuspendLayout();
            splitContainer.SuspendLayout();
            pnlLogTop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvEvents).BeginInit();
            pnlFooter.SuspendLayout();
            SuspendLayout();

            // ── pnlHeader ────────────────────────────────────────────────────────────
            pnlHeader.Controls.Add(btnConnectDisconnect);
            pnlHeader.Dock     = System.Windows.Forms.DockStyle.Top;
            pnlHeader.Size     = new System.Drawing.Size(920, 48);
            pnlHeader.Name     = "pnlHeader";
            pnlHeader.TabIndex = 0;

            btnConnectDisconnect.Anchor   = System.Windows.Forms.AnchorStyles.Right | System.Windows.Forms.AnchorStyles.Top;
            btnConnectDisconnect.Location = new System.Drawing.Point(820, 10);
            btnConnectDisconnect.Name     = "btnConnectDisconnect";
            btnConnectDisconnect.Size     = new System.Drawing.Size(92, 28);
            btnConnectDisconnect.TabIndex = 0;
            btnConnectDisconnect.Text     = "Connect";
            btnConnectDisconnect.UseVisualStyleBackColor = true;
            btnConnectDisconnect.Click   += new System.EventHandler(btnConnectDisconnect_Click);

            // ── splitContainer ────────────────────────────────────────────────────────
            splitContainer.Dock             = System.Windows.Forms.DockStyle.Fill;
            splitContainer.Name             = "splitContainer";
            splitContainer.SplitterDistance = 500;
            splitContainer.TabIndex         = 1;

            splitContainer.Panel1.Controls.Add(rtbLog);
            splitContainer.Panel1.Controls.Add(pnlLogTop);
            splitContainer.Panel2.Controls.Add(dgvEvents);

            pnlLogTop.Controls.Add(btnClearLog);
            pnlLogTop.Dock     = System.Windows.Forms.DockStyle.Top;
            pnlLogTop.Height   = 36;
            pnlLogTop.Name     = "pnlLogTop";
            pnlLogTop.TabIndex = 0;

            btnClearLog.Anchor   = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Top;
            btnClearLog.Location = new System.Drawing.Point(8, 4);
            btnClearLog.Name     = "btnClearLog";
            btnClearLog.Size     = new System.Drawing.Size(90, 28);
            btnClearLog.TabIndex = 0;
            btnClearLog.Text     = "Clear Log";
            btnClearLog.UseVisualStyleBackColor = true;
            btnClearLog.Click   += new System.EventHandler(btnClearLog_Click);

            rtbLog.BackColor  = System.Drawing.Color.Black;
            rtbLog.Dock       = System.Windows.Forms.DockStyle.Fill;
            rtbLog.Font       = new System.Drawing.Font("Consolas", 9F);
            rtbLog.ForeColor  = System.Drawing.Color.Lime;
            rtbLog.Name       = "rtbLog";
            rtbLog.ReadOnly   = true;
            rtbLog.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            rtbLog.TabIndex   = 1;
            rtbLog.Text       = "";
            rtbLog.WordWrap   = false;

            dgvEvents.AllowUserToAddRows     = false;
            dgvEvents.AllowUserToDeleteRows  = false;
            dgvEvents.AutoSizeColumnsMode    = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dgvEvents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvEvents.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { colName, colCategory, colLog });
            dgvEvents.Dock                   = System.Windows.Forms.DockStyle.Fill;
            dgvEvents.MultiSelect            = false;
            dgvEvents.Name                   = "dgvEvents";
            dgvEvents.RowHeadersVisible      = false;
            dgvEvents.SelectionMode          = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            dgvEvents.TabIndex               = 0;
            dgvEvents.CurrentCellDirtyStateChanged += new System.EventHandler(dgvEvents_CurrentCellDirtyStateChanged);
            dgvEvents.CellValueChanged       += new System.Windows.Forms.DataGridViewCellEventHandler(dgvEvents_CellValueChanged);

            colName.FillWeight = 50;
            colName.HeaderText = "Name";
            colName.Name       = "colName";
            colName.ReadOnly   = true;

            colCategory.FillWeight = 30;
            colCategory.HeaderText = "Category";
            colCategory.Name       = "colCategory";
            colCategory.ReadOnly   = true;

            colLog.FillWeight = 20;
            colLog.HeaderText = "Log";
            colLog.Name       = "colLog";
            colLog.SortMode   = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;

            // ── pnlFooter ────────────────────────────────────────────────────────────
            pnlFooter.Controls.Add(btnSelectOutputFile);
            pnlFooter.Controls.Add(lblJsonWriteMethod);
            pnlFooter.Controls.Add(cbxJsonWriteMethod);
            pnlFooter.Controls.Add(txtOutputFilePath);
            pnlFooter.Controls.Add(btnSaveAsJson);
            pnlFooter.Dock     = System.Windows.Forms.DockStyle.Bottom;
            pnlFooter.Size     = new System.Drawing.Size(920, 44);
            pnlFooter.Name     = "pnlFooter";
            pnlFooter.TabIndex = 2;

            btnSelectOutputFile.Location = new System.Drawing.Point(8, 7);
            btnSelectOutputFile.Name     = "btnSelectOutputFile";
            btnSelectOutputFile.Size     = new System.Drawing.Size(90, 28);
            btnSelectOutputFile.TabIndex = 0;
            btnSelectOutputFile.Text     = "Select...";
            btnSelectOutputFile.UseVisualStyleBackColor = true;
            btnSelectOutputFile.Click   += new System.EventHandler(btnSelectOutputFile_Click);
            lblJsonWriteMethod.AutoSize = true;
            lblJsonWriteMethod.Location = new System.Drawing.Point(104, 13);
            lblJsonWriteMethod.Name     = "lblJsonWriteMethod";
            lblJsonWriteMethod.Text     = "JSON Write Method:";
            lblJsonWriteMethod.TabIndex = 1;

            cbxJsonWriteMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cbxJsonWriteMethod.Location = new System.Drawing.Point(242, 8);
            cbxJsonWriteMethod.Name     = "cbxJsonWriteMethod";
            cbxJsonWriteMethod.Size     = new System.Drawing.Size(140, 23);
            cbxJsonWriteMethod.TabIndex = 2;
            txtOutputFilePath.Anchor   = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right | System.Windows.Forms.AnchorStyles.Top;
            txtOutputFilePath.Location = new System.Drawing.Point(390, 10);
            txtOutputFilePath.Name     = "txtOutputFilePath";
            txtOutputFilePath.Size     = new System.Drawing.Size(414, 23);
            txtOutputFilePath.TabIndex = 3;

            btnSaveAsJson.Anchor   = System.Windows.Forms.AnchorStyles.Right | System.Windows.Forms.AnchorStyles.Top;
            btnSaveAsJson.Location = new System.Drawing.Point(810, 7);
            btnSaveAsJson.Name     = "btnSaveAsJson";
            btnSaveAsJson.Size     = new System.Drawing.Size(102, 28);
            btnSaveAsJson.TabIndex = 2;
            btnSaveAsJson.Text     = "Save As JSON";
            btnSaveAsJson.UseVisualStyleBackColor = true;
            btnSaveAsJson.Click   += new System.EventHandler(btnSaveAsJson_Click);

            // ── MsPowerPoint_Watcher_Form ─────────────────────────────────────────────
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize          = new System.Drawing.Size(920, 640);
            Controls.Add(splitContainer);
            Controls.Add(pnlHeader);
            Controls.Add(pnlFooter);
            MinimumSize         = new System.Drawing.Size(700, 480);
            Name                = "MsPowerPoint_Watcher_Form";
            StartPosition       = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text                = "Microsoft PowerPoint - Watcher";
            Load                += new System.EventHandler(MsPowerPoint_Watcher_Form_Load);
            FormClosing         += new System.Windows.Forms.FormClosingEventHandler(MsPowerPoint_Watcher_Form_FormClosing);

            pnlHeader.ResumeLayout(false);
            pnlHeader.PerformLayout();
            splitContainer.Panel1.ResumeLayout(false);
            splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer).EndInit();
            splitContainer.ResumeLayout(false);
            pnlLogTop.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvEvents).EndInit();
            pnlFooter.ResumeLayout(false);
            pnlFooter.PerformLayout();
            ResumeLayout(false);
        }

        private System.Windows.Forms.Panel               pnlHeader;
        private System.Windows.Forms.Button              btnConnectDisconnect;
        private System.Windows.Forms.SplitContainer      splitContainer;
        private System.Windows.Forms.Panel               pnlLogTop;
        private System.Windows.Forms.Button              btnClearLog;
        private System.Windows.Forms.RichTextBox         rtbLog;
        private System.Windows.Forms.DataGridView        dgvEvents;
        private System.Windows.Forms.DataGridViewTextBoxColumn  colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn  colCategory;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colLog;
        private System.Windows.Forms.Panel               pnlFooter;
        private System.Windows.Forms.Button              btnSelectOutputFile;
        private System.Windows.Forms.Label               lblJsonWriteMethod;
        private System.Windows.Forms.ComboBox            cbxJsonWriteMethod;
        private System.Windows.Forms.TextBox             txtOutputFilePath;
        private System.Windows.Forms.Button              btnSaveAsJson;
    }
}
