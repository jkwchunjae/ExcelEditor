using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public class Metadata
    {
        [JsonIgnore]
        public bool Dirty { get; set; }

        [JsonProperty("columns")]
        public List<ColumnSetting> Columns { get; set; }
    }
}
