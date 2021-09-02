using EeCommon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeJson
{
    public class JsonArrayElement : JsonBaseElement, IArrayElement
    {
        public int Length => JArray.Count;
        public bool Any => Length != 0;
        public bool Empty => Length == 0;


        private JArray JArray;

        public JsonArrayElement(JArray array)
            : base(array)
        {
            JArray = array;
        }
        public JsonArrayElement(JsonBaseElement baseElement)
            : base(baseElement.Token)
        {
            JArray = (JArray)baseElement.Token;
        }
    }
}
