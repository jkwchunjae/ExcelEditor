using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public class Metadata
    {
        [JsonProperty("columns")]
        public List<ColumnSetting> Columns { get; set; }
    }
}
