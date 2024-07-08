using Newtonsoft.Json;


namespace ToolsV2Classes
{
    public class Fields
    {
        [JsonProperty("SKU")]
        public string SKU { get; set; }

        [JsonProperty("Unit Price")]
        public string UnitPrice { get; set; }

        [JsonProperty("Generic Name ")]
        public string GenericName  { get; set; }

        [JsonProperty("Approval")]
        public string approvalStatus { get; set; }

        [JsonProperty("Subcategory")]
        public string Subcategory { get; set; }

    }
}
