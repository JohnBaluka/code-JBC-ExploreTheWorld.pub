namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public class MsOfficeEvent_Record
    {
        public string Name       { get; set; } = "";
        public string Category   { get; set; } = "";
        public bool   Word       { get; set; }
        public bool   Excel      { get; set; }
        public bool   PowerPoint { get; set; }

        /// <summary>Mutable — user can toggle per session via the events DataGridView.</summary>
        public bool Log { get; set; } = true;
    }
}
