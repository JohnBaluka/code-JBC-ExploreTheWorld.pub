namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    partial class ExportLog_Form
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
            rtbLog = new System.Windows.Forms.RichTextBox();
            pnlBottom = new System.Windows.Forms.Panel();
            btnRunExport = new System.Windows.Forms.Button();
            lblStatus = new System.Windows.Forms.Label();
            pnlBottom.SuspendLayout();
            SuspendLayout();

            // rtbLog
            rtbLog.BackColor = System.Drawing.Color.Black;
            rtbLog.Dock = System.Windows.Forms.DockStyle.Fill;
            rtbLog.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            rtbLog.ForeColor = System.Drawing.Color.Lime;
            rtbLog.Location = new System.Drawing.Point(0, 0);
            rtbLog.Name = "rtbLog";
            rtbLog.ReadOnly = true;
            rtbLog.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            rtbLog.Size = new System.Drawing.Size(900, 520);
            rtbLog.TabIndex = 0;
            rtbLog.Text = "";
            rtbLog.WordWrap = false;

            // pnlBottom
            pnlBottom.Controls.Add(btnRunExport);
            pnlBottom.Controls.Add(lblStatus);
            pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            pnlBottom.Location = new System.Drawing.Point(0, 520);
            pnlBottom.Name = "pnlBottom";
            pnlBottom.Size = new System.Drawing.Size(900, 40);
            pnlBottom.TabIndex = 1;

            // lblStatus
            lblStatus.AutoSize = true;
            lblStatus.Location = new System.Drawing.Point(8, 12);
            lblStatus.Name = "lblStatus";
            lblStatus.Text = "Click \"Run Export\" to choose a save location and begin.";

            // btnRunExport
            btnRunExport.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            btnRunExport.Location = new System.Drawing.Point(780, 8);
            btnRunExport.Name = "btnRunExport";
            btnRunExport.Size = new System.Drawing.Size(112, 26);
            btnRunExport.TabIndex = 0;
            btnRunExport.Text = "Run Export";
            btnRunExport.UseVisualStyleBackColor = true;
            btnRunExport.Click += new System.EventHandler(btnRunExport_Click);

            // ExportLog_Form
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(900, 560);
            Controls.Add(rtbLog);
            Controls.Add(pnlBottom);
            MinimumSize = new System.Drawing.Size(640, 400);
            Name = "ExportLog_Form";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Export Log";
            pnlBottom.ResumeLayout(false);
            pnlBottom.PerformLayout();
            ResumeLayout(false);
        }

        private System.Windows.Forms.RichTextBox rtbLog;
        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Button btnRunExport;
        private System.Windows.Forms.Label lblStatus;
    }
}
