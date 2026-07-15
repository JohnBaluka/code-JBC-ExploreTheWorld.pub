using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Shape : Shape_Fields
    {
        // Entities
        public AnimationSettings? AnimationSettings { get; set; }
        public FillFormat? Fill { get; set; }
        public LineFormat? Line { get; set; }
        public PictureFormat? PictureFormat { get; set; }
        public PlaceholderFormat? PlaceholderFormat { get; set; }
        public TextFrame? TextFrame { get; set; }
        public TextFrame2? TextFrame2 { get; set; }
        public ImageBlob? Image { get; set; }

        // Lists
        public List<ActionSetting>? ActionSettings { get; set; }
        public List<Tag>? Tags { get; set; }
        public List<Shape>? GroupItems { get; set; }

        // Enums
        [JsonIgnore]
        public ShapeType_Enum Type_Enum
        {
            get
            {
                return (ShapeType_Enum)Type;
            }
            set
            {
                Type = (int)value;
            }
        }

        [JsonIgnore]
        public AutoShapeType_Enum AutoShapeType_Enum
        {
            get
            {
                return (AutoShapeType_Enum)AutoShapeType;
            }
            set
            {
                AutoShapeType = (int)value;
            }
        }

        [JsonIgnore]
        public string CodeName
        {
            get
            {
                if (Name == null) return string.Empty;

                return Name.Replace(" ", "_").Replace("!", "_");
            }
        }

        public bool IsVisible()
        {
            if (!Visible.HasValue) return false;

            return TriStateExtensions.IsTrue(Visible.Value);
        }

        public bool IsLineVisible()
        {
            if (Line == null || !Line.Visible.HasValue) return false;

            return TriStateExtensions.IsTrue(Line.Visible.Value);
        }

        public bool IsNarrationAsBoolean()
        {
            if (!IsNarration.HasValue) return false;

            return TriStateExtensions.IsTrue(IsNarration.Value);
        }

        public bool IsUsedIfPlaceHolder()
        {
            if (Type_Enum != ShapeType_Enum.msoPlaceholder) return true;

            if (HasChart.HasValue && TriStateExtensions.IsTrue(HasChart.Value)) return true;
            if (HasInkXML.HasValue && TriStateExtensions.IsTrue(HasInkXML.Value)) return true;
            if (HasSectionZoom.HasValue && TriStateExtensions.IsTrue(HasSectionZoom.Value)) return true;
            if (HasSmartArt.HasValue && TriStateExtensions.IsTrue(HasSmartArt.Value)) return true;
            if (HasTable.HasValue && TriStateExtensions.IsTrue(HasTable.Value)) return true;

            if (!string.IsNullOrEmpty(GetText())) return true;

            if (PlaceholderFormat != null
                && PlaceholderFormat.ContainedType.HasValue
                && PlaceholderFormat.ContainedType.Value != 1) return true;

            return false;
        }

        public override string ToString()
        {
            string s = Name + " : " + Type_Enum.ToString();

            if (Type_Enum == ShapeType_Enum.msoAutoShape)
            {
                s += " : " + AutoShapeType_Enum.ToString();
            }

            s += " : " + Title;

            return s;
        }

        public string GetText()
        {
            if (TextFrame2 == null) return string.Empty;
            if (TextFrame2.TextRange2 == null) return string.Empty;
            if (TextFrame2.TextRange2.Text == null) return string.Empty;

            return TextFrame2.TextRange2.Text;
        }
    }
}
