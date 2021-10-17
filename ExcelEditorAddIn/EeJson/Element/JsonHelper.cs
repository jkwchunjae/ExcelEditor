using EeCommon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeJson
{
    public class JsonHelper : IHelper
    {
        public T Deserialize<T>(string text)
        {
            var obj = JsonConvert.DeserializeObject<T>(text);
            return obj;
        }

        public string Serialize<T>(T obj)
        {
            var text = JsonConvert.SerializeObject(obj, Formatting.Indented);
            return text;
        }

        public string MetadataFilePath(string filePath)
        {
            var ext = Path.GetExtension(filePath);
            var filePathWithoutExt = filePath.Substring(0, filePath.Length - ext.Length);

            var metadataPath = $@"{filePathWithoutExt}.meta.json";
            return metadataPath;
        }
    }
}
