using Newtonsoft.Json;

namespace CallContent.Models
{
    public class WebPartPatch
    {
        [JsonProperty("@odata.type")]
        public string ODataType { get; set; }

        public string id { get; set; }
        public string innerHtml { get; set; }
    }
}
