using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // The fully resolved character formatting of one text run, after walking the
    // placeholder/list-style inheritance chain. TextRange.Font aggregates these per
    // shape (equal values pass through, differing values become the COM mixed markers).
    internal sealed class OpenXmlEffectiveRunFormat
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }

        // COM range getters report mixed when a run turns bold/italic off explicitly
        // while sibling runs merely inherit the off state.
        public bool BoldExplicitOff { get; set; }
        public bool ItalicExplicitOff { get; set; }

        // The last paragraph's mark formatting participates in the Color aggregation
        // only (it carries no visible character).
        public bool ColorOnly { get; set; }
        public int UnderlineStyle { get; set; }
        public int Caps { get; set; }
        public int Strike { get; set; }
        public double SizePoints { get; set; } = 18;
        public double BaselineFraction { get; set; }
        public double SpacingPoints { get; set; }
        public double KerningPoints { get; set; } = 12;
        public string LatinName { get; set; } = string.Empty;
        public string FarEastName { get; set; } = "+mn-ea";
        public string ComplexScriptName { get; set; } = "+mn-cs";
        public PP.ColorFormat? Color { get; set; }
    }
}
