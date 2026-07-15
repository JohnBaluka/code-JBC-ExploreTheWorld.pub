using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Document : Document_Fields
    {
        // Lists
        public List<BuiltInDocumentProperty>? BuiltInDocumentProperties { get; set; }
        public List<CustomDocumentProperty>? CustomDocumentProperties { get; set; }
        public List<Bookmark>? Bookmarks { get; set; }
        public List<Comment>? Comments { get; set; }
        public List<ContentControl>? ContentControls { get; set; }
        public List<Endnote>? Endnotes { get; set; }
        public List<Field>? Fields { get; set; }
        public List<Footnote>? Footnotes { get; set; }
        public List<InlineShape>? InlineShapes { get; set; }
        public List<Shape>? Shapes { get; set; }
        public List<Paragraph>? Paragraphs { get; set; }
        public List<Section>? Sections { get; set; }
        public List<Style>? Styles { get; set; }
        public List<Table>? Tables { get; set; }
        public List<Variable>? Variables { get; set; }
    }
}
