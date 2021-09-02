using EeCommon;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;

namespace EeJson
{
    public class JsonArrayDocument : JsonBaseDocument, IArrayDocument
    {
        public int Length => Doc?.RootElement.GetArrayLength() ?? 0;
        public bool Any => Length != 0;
        public bool Empty => Length == 0;

        public JsonArrayDocument(string jsonText)
            : base(jsonText)
        {
        }
        public JsonArrayDocument(JsonBaseDocument baseDocument)
            : base(baseDocument.Doc)
        {
        }
    }
}
