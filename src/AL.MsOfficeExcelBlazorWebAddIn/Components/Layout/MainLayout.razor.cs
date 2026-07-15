namespace ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Components.Layout
{
    public partial class MainLayout
    {
        public string TrademarkMessage1 { get; set; } = "Copyright © " + @DateTime.Now.Year + " John Baluka.";
        public string TrademarkMessage2 { get; set; } = "All rights reserved.";
        public string FrameworkDescription { get; set; } = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
    }
}
