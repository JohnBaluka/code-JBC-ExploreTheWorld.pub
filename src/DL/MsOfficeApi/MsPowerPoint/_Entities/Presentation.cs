using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Presentation : Presentation_Fields
    {
        // Entities
        public SlideMaster? SlideMaster { get; set; }
        public SlideMaster? TitleMaster { get; set; }
        public SlideMaster? HandoutMaster { get; set; }
        public SlideMaster? NotesMaster { get; set; }
        public Coauthoring? Coauthoring { get; set; }
        public PageSetup? PageSetup { get; set; }
        public PrintOptions? PrintOptions { get; set; }
        public SlideShowSettings? SlideShowSettings { get; set; }

        // Lists
        public List<BuiltInDocumentProperty>? BuiltInDocumentProperties { get; set; }
        public List<CustomDocumentProperty>? CustomDocumentProperties { get; set; }
        public List<ContentTypeProperty>? ContentTypeProperties { get; set; }
        public List<ColorScheme>? ColorSchemes { get; set; }
        public List<CustomXmlPart>? CustomXmlParts { get; set; }
        public List<Design>? Designs { get; set; }
        public List<long>? ExtraColors { get; set; }
        public List<PresentationFont>? Fonts { get; set; }
        public List<Guide>? Guides { get; set; }
        public List<SectionProperty>? SectionProperties { get; set; }
        public List<Slide>? Slides { get; set; }
        public List<Tag>? Tags { get; set; }

        [JsonIgnore]
        public DateTime? Created { get; set; }

        [JsonIgnore]
        public DateTime? Modified { get; set; }
    }
}
