namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    partial class Main_Form
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
            btnCountriesNowSpace = new System.Windows.Forms.Button();
            btnWatcher           = new System.Windows.Forms.Button();
            cmsWatcher           = new System.Windows.Forms.ContextMenuStrip();
            mnuWatcherWord       = new System.Windows.Forms.ToolStripMenuItem();
            mnuWatcherExcel      = new System.Windows.Forms.ToolStripMenuItem();
            mnuWatcherPowerPoint = new System.Windows.Forms.ToolStripMenuItem();
            cmsWatcher.SuspendLayout();
            SuspendLayout();

            // btnCountriesNowSpace
            btnCountriesNowSpace.Location = new System.Drawing.Point(40, 40);
            btnCountriesNowSpace.Name = "btnCountriesNowSpace";
            btnCountriesNowSpace.Size = new System.Drawing.Size(240, 48);
            btnCountriesNowSpace.TabIndex = 0;
            btnCountriesNowSpace.Text = "Countries Now API";
            btnCountriesNowSpace.UseVisualStyleBackColor = true;
            btnCountriesNowSpace.Click += new System.EventHandler(btnCountriesNowSpace_Click);

            // cmsWatcher
            cmsWatcher.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                mnuWatcherWord, mnuWatcherExcel, mnuWatcherPowerPoint });
            cmsWatcher.Name = "cmsWatcher";

            // mnuWatcherWord
            mnuWatcherWord.Name   = "mnuWatcherWord";
            mnuWatcherWord.Text   = "Word";
            mnuWatcherWord.Click += new System.EventHandler(mnuWatcherWord_Click);

            // mnuWatcherExcel
            mnuWatcherExcel.Name   = "mnuWatcherExcel";
            mnuWatcherExcel.Text   = "Excel";
            mnuWatcherExcel.Click += new System.EventHandler(mnuWatcherExcel_Click);

            // mnuWatcherPowerPoint
            mnuWatcherPowerPoint.Name   = "mnuWatcherPowerPoint";
            mnuWatcherPowerPoint.Text   = "PowerPoint";
            mnuWatcherPowerPoint.Click += new System.EventHandler(mnuWatcherPowerPoint_Click);

            // btnWatcher
            btnWatcher.Location = new System.Drawing.Point(40, 104);
            btnWatcher.Name     = "btnWatcher";
            btnWatcher.Size     = new System.Drawing.Size(240, 48);
            btnWatcher.TabIndex = 1;
            btnWatcher.Text     = "Watcher \u25bc";
            btnWatcher.UseVisualStyleBackColor = true;
            btnWatcher.Click   += new System.EventHandler(btnWatcher_Click);

            // Main_Form
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(320, 200);
            Controls.Add(btnCountriesNowSpace);
            Controls.Add(btnWatcher);
            Name = "Main_Form";
            Text = "ETW WinFormApp._netF";
            cmsWatcher.ResumeLayout(false);
            ResumeLayout(false);
        }

        private System.Windows.Forms.Button           btnCountriesNowSpace;
        private System.Windows.Forms.Button           btnWatcher;
        private System.Windows.Forms.ContextMenuStrip cmsWatcher;
        private System.Windows.Forms.ToolStripMenuItem mnuWatcherWord;
        private System.Windows.Forms.ToolStripMenuItem mnuWatcherExcel;
        private System.Windows.Forms.ToolStripMenuItem mnuWatcherPowerPoint;
    }
}
