using System;
using System.Windows.Forms;
using JBC.ExploreTheWorld.BL;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class Main_Form : Form
    {
        private readonly CountriesNowSpaceManager__Service _countriesNowManager;

        public Main_Form(
            CountriesNowSpaceManager__Service countriesNowManager)
        {
            InitializeComponent();
            _countriesNowManager = countriesNowManager;
        }

        private void btnCountriesNowSpace_Click(object sender, EventArgs e)
        {
            new CountriesNowSpace_Form(_countriesNowManager).Show();
        }

        private void btnWatcher_Click(object sender, EventArgs e)
        {
            cmsWatcher.Show(btnWatcher, 0, btnWatcher.Height);
        }

        private void mnuWatcherWord_Click(object sender, EventArgs e)
        {
            new MsWord_Watcher_Form().Show();
        }

        private void mnuWatcherExcel_Click(object sender, EventArgs e)
        {
            new MsExcel_Watcher_Form().Show();
        }

        private void mnuWatcherPowerPoint_Click(object sender, EventArgs e)
        {
            new MsPowerPoint_Watcher_Form().Show();
        }
    }
}
