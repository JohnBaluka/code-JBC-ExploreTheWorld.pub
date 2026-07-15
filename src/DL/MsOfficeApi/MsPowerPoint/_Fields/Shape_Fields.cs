using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Shape_Fields
    {
        public long? Id { get; set; }
        public string? Name { get; set; }

        public int? Type { get; set; }
        public int? AutoShapeType { get; set; }
        public string? AlternativeText { get; set; }
        public string? Title { get; set; }

        public string? Left { get; set; }
        public string? Top { get; set; }
        public string? Width { get; set; }
        public string? Height { get; set; }
        public string? Rotation { get; set; }
        public int? ZOrderPosition { get; set; }
        public int? Visible { get; set; }

        public int? BackgroundStyle { get; set; }
        public int? BlackWhiteMode { get; set; }
        public int? Child { get; set; }
        public long? ConnectionSiteCount { get; set; }
        public int? Connector { get; set; }
        public long? Creator { get; set; }
        public int? Decorative { get; set; }
        public int? GraphicStyle { get; set; }
        public int? HasChart { get; set; }
        public int? HasInkXML { get; set; }
        public int? HasSectionZoom { get; set; }
        public int? HasSmartArt { get; set; }
        public int? HasTable { get; set; }
        public int? HasTextFrame { get; set; }
        public int? HorizontalFlip { get; set; }
        public string? InkXML { get; set; }
        public int? IsNarration { get; set; }
        public int? LockAspectRatio { get; set; }
        public int? ShapeStyle { get; set; }
        public int? VerticalFlip { get; set; }
        public string? Vertices { get; set; }
        public int? MediaType { get; set; }

    }
}
