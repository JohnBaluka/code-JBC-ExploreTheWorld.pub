namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    partial class ExploreTheWorld_Form
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

            // blazorWebView
            blazorWebView.Dock     = System.Windows.Forms.DockStyle.Fill;
            blazorWebView.Name     = "blazorWebView";
            blazorWebView.TabIndex = 0;

            // ExploreTheWorld_Form
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize          = new System.Drawing.Size(1280, 800);
            Controls.Add(blazorWebView);
            Name = "ExploreTheWorld_Form";
            Text = "ETW WinFormApp";
            ResumeLayout(false);
        }

        private Microsoft.AspNetCore.Components.WebView.WindowsForms.BlazorWebView blazorWebView;
    }
}
