using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeCommon
{
    public abstract class ValueConvertException : Exception
    {
        public object Value { get; private set; }
        public object Value2 { get; private set; }
        public ValueConvertException(object value, object value2)
        {
            Value = value;
            Value2 = value2;
        }
    }

    public class RequireNumberException : ValueConvertException
    {
        public RequireNumberException(object value, object value2)
            : base(value, value2)
        {
        }
    }

    public class RequireBooleanException : ValueConvertException
    {
        public RequireBooleanException(object value, object value2)
            : base(value, value2)
        {
        }
    }
}
