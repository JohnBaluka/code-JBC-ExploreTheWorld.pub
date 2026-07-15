using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    // Result of getPresentationInfo() in wwwroot/js/presentation-info.js.
    public class PowerPointPresentationInfo_Row
    {
        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        [JsonPropertyName("slideCount")]
        public int SlideCount { get; set; }

        [JsonPropertyName("slideWidth")]
        public double SlideWidth { get; set; }

        [JsonPropertyName("slideHeight")]
        public double SlideHeight { get; set; }

        [JsonPropertyName("author")]
        public string Author { get; set; } = string.Empty;

        [JsonPropertyName("errorMessage")]
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
