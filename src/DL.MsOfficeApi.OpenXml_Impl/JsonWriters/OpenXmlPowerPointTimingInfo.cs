using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Animation build data extracted from a slide's p:timing tree. AnimationOrder is
    // the 1-based position of the shape's first effect node in document order;
    // OnClick reflects the effect's trigger node type.
    internal sealed class OpenXmlPowerPointTimingInfo
    {
        private readonly Dictionary<uint, (long Order, bool OnClick)> _targets = new Dictionary<uint, (long, bool)>();
        private readonly HashSet<uint> _narrationTargets = new HashSet<uint>();

        public int ClickCount { get; private set; }

        public bool IsNarrationTarget(uint shapeId)
        {
            return _narrationTargets.Contains(shapeId);
        }

        public bool TryGetTarget(uint shapeId, out long order, out bool onClick)
        {
            if (_targets.TryGetValue(shapeId, out var info))
            {
                order = info.Order;
                onClick = info.OnClick;
                return true;
            }

            order = 0;
            onClick = false;
            return false;
        }

        public static OpenXmlPowerPointTimingInfo Parse(P.Timing? timing)
        {
            var result = new OpenXmlPowerPointTimingInfo();
            if (timing == null) return result;

            // Narration audio is marked with isNarration="1" on p:audio timing nodes.
            foreach (var audio in timing.Descendants<P.Audio>())
            {
                if (audio.IsNarration?.Value != true) continue;

                foreach (var target in audio.Descendants<P.ShapeTarget>())
                {
                    if (target.ShapeId?.Value is string narrationSpid && uint.TryParse(narrationSpid, out uint narrationId))
                    {
                        result._narrationTargets.Add(narrationId);
                    }
                }
            }

            long effectIndex = 0;
            foreach (var node in timing.Descendants<P.CommonTimeNode>())
            {
                string? nodeType = node.NodeType?.InnerText;
                if (nodeType != "clickEffect" && nodeType != "withEffect" && nodeType != "afterEffect") continue;

                effectIndex++;
                if (nodeType == "clickEffect") result.ClickCount++;

                var target = node.Descendants<P.ShapeTarget>().FirstOrDefault();
                if (target?.ShapeId?.Value is string spid
                    && uint.TryParse(spid, out uint shapeId)
                    && !result._targets.ContainsKey(shapeId))
                {
                    result._targets.Add(shapeId, (effectIndex, nodeType == "clickEffect"));
                }
            }

            return result;
        }
    }
}
