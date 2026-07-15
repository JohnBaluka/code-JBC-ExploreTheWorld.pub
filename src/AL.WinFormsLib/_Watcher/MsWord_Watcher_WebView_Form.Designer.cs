namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    partial class MsWord_Watcher_WebView_Form
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
            blazorWebView = new Microsoft.AspNetCore.Components.WebView.WindowsForms.BlazorWebView();
            SuspendLayout();

            blazorWebView.Dock     = System.Windows.Forms.DockStyle.Fill;
            blazorWebView.Name     = "blazorWebView";
            blazorWebView.TabIndex = 0;

            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize          = new System.Drawing.Size(1200, 800);
            Controls.Add(blazorWebView);
            Name  = "MsWord_Watcher_WebView_Form";
            Text  = "Microsoft Word - Watcher (Web)";
            ResumeLayout(false);
        }

        private Microsoft.AspNetCore.Components.WebView.WindowsForms.BlazorWebView blazorWebView;
    }
}
