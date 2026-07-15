using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Presentation_Fields
    {
        public string? Name { get; set; }
        public string? Path { get; set; }
        public string? FullName { get; set; }

        public bool? AutoSaveOn { get; set; }
        public bool? ChartDataPointTrack { get; set; }
        public int? CreateVideoStatus { get; set; }
        public int? DefaultLanguageID { get; set; }
        public int? DisplayComments { get; set; }
        public string? EncryptionProvider { get; set; }
        public int? EnvelopeVisible { get; set; }
        public int? FarEastLineBreakLanguage { get; set; }
        public int? FarEastLineBreakLevel { get; set; }
        public bool? Final { get; set; }
        public string? GridDistance { get; set; }
        public int? HasHandoutMaster { get; set; }
        public int? HasNotesMaster { get; set; }
        public int? HasTitleMaster { get; set; }
        public int? HasVBProject { get; set; }
        public bool? InMergeMode { get; set; }
        public int? LayoutDirection { get; set; }
        public string? NoLineBreakAfter { get; set; }
        public string? NoLineBreakBefore { get; set; }
        public string? Password { get; set; }
        public string? PasswordEncryptionAlgorithm { get; set; }
        public bool? PasswordEncryptionFileProperties { get; set; }
        public long? PasswordEncryptionKeyLength { get; set; }
        public string? PasswordEncryptionProvider { get; set; }
        public int? ReadOnly { get; set; }
        public bool? ReadOnlyRecommended { get; set; }
        public int? RemovePersonalInformation { get; set; }
        public int? Saved { get; set; }
        public int? SnapToGrid { get; set; }
        public string? TemplateName { get; set; }
        public int? VBASigned { get; set; }
        public string? WritePassword { get; set; }
    }
}
