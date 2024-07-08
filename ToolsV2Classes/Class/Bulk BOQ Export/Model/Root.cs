using Newtonsoft.Json;
using System.Collections.Generic;

namespace ToolsV2Classes
{
    public class Root
    {
        [JsonProperty("records")]
        public List<Record> records { get; set; }

        [JsonProperty("offset")]
        public string Offset { get; set; }
    }
}
