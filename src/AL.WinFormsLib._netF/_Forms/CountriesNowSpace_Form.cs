using System.Windows.Forms;
using JBC.ExploreTheWorld.BL;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class CountriesNowSpace_Form : Form
    {
        private readonly CountriesNowSpace_UserControl _userControl;

        public CountriesNowSpace_Form(CountriesNowSpaceManager__Service manager)
        {
            InitializeComponent();
            _userControl = new CountriesNowSpace_UserControl(manager) { Dock = DockStyle.Fill };
            Controls.Add(_userControl);
        }

        /// <summary>
        /// Switches the hosted Countries control into add-in export mode (export type locked to the
        /// host Office app, library locked to NetOffice, export into a new in-app document).
        /// </summary>
        public void EnableAddinExportMode(string hostType) => _userControl.EnableAddinExportMode(hostType);
    }
}
