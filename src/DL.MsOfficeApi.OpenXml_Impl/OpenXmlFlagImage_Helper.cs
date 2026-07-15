using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using P = DocumentFormat.OpenXml.Presentation;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl
{
    /// <summary>
    /// Builds the DrawingML fragments used to embed PNG flag images in Word tables,
    /// Excel worksheets, and PowerPoint slides. The image parts themselves are added by
    /// the callers (each host part type has its own <c>AddImagePart</c>).
    /// </summary>
    internal static class OpenXmlFlagImage_Helper
    {
        /// <summary>Inline picture for a Word table cell.</summary>
        public static W.Drawing CreateWordInlineDrawing(string relationshipId, uint drawingId, string name, long cx, long cy)
            => new(new DW.Inline(
                       new DW.Extent { Cx = cx, Cy = cy },
                       new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                       new DW.DocProperties { Id = drawingId, Name = name },
                       new DW.NonVisualGraphicFrameDrawingProperties(
                           new A.GraphicFrameLocks { NoChangeAspect = true }),
                       new A.Graphic(
                           new A.GraphicData(CreatePicture(relationshipId, drawingId, name, cx, cy))
                           { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                   {
                       DistanceFromTop    = 0U,
                       DistanceFromBottom = 0U,
                       DistanceFromLeft   = 0U,
                       DistanceFromRight  = 0U
                   });

        /// <summary>Picture shape for a PowerPoint slide shape tree.</summary>
        public static P.Picture CreatePresentationPicture(
            string relationshipId, uint shapeId, string name, long x, long y, long cx, long cy)
            => new(new P.NonVisualPictureProperties(
                       new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                       new P.NonVisualPictureDrawingProperties(
                           new A.PictureLocks { NoChangeAspect = true }),
                       new P.ApplicationNonVisualDrawingProperties()),
                   new P.BlipFill(
                       new A.Blip { Embed = relationshipId },
                       new A.Stretch(new A.FillRectangle())),
                   new P.ShapeProperties(
                       new A.Transform2D(
                           new A.Offset  { X = x, Y = y },
                           new A.Extents { Cx = cx, Cy = cy }),
                       new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        /// <summary>One-cell-anchored picture for an Excel worksheet drawing.</summary>
        public static Xdr.OneCellAnchor CreateSpreadsheetAnchor(
            string relationshipId, uint drawingId, string name, int columnIndex, int rowIndex, long cx, long cy)
            => new(new Xdr.FromMarker(
                       new Xdr.ColumnId(columnIndex.ToString()),
                       new Xdr.ColumnOffset("19050"),
                       new Xdr.RowId(rowIndex.ToString()),
                       new Xdr.RowOffset("19050")),
                   new Xdr.Extent { Cx = cx, Cy = cy },
                   new Xdr.Picture(
                       new Xdr.NonVisualPictureProperties(
                           new Xdr.NonVisualDrawingProperties { Id = drawingId, Name = name },
                           new Xdr.NonVisualPictureDrawingProperties(
                               new A.PictureLocks { NoChangeAspect = true })),
                       new Xdr.BlipFill(
                           new A.Blip { Embed = relationshipId },
                           new A.Stretch(new A.FillRectangle())),
                       new Xdr.ShapeProperties(
                           new A.Transform2D(
                               new A.Offset  { X = 0L, Y = 0L },
                               new A.Extents { Cx = cx, Cy = cy }),
                           new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                   new Xdr.ClientData());

        private static PIC.Picture CreatePicture(string relationshipId, uint id, string name, long cx, long cy)
            => new(new PIC.NonVisualPictureProperties(
                       new PIC.NonVisualDrawingProperties { Id = id, Name = name },
                       new PIC.NonVisualPictureDrawingProperties()),
                   new PIC.BlipFill(
                       new A.Blip { Embed = relationshipId },
                       new A.Stretch(new A.FillRectangle())),
                   new PIC.ShapeProperties(
                       new A.Transform2D(
                           new A.Offset  { X = 0L, Y = 0L },
                           new A.Extents { Cx = cx, Cy = cy }),
                       new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }
}
