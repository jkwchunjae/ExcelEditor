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
        public ValueConvertException(object value, object value2, string message)
            : base(message)
        {
            Value = value;
            Value2 = value2;
        }
    }

    public class RequireNumberException : ValueConvertException
    {
        static readonly string ErrorMessage = "숫자 형식을 입력해야 합니다.";
        public RequireNumberException(object value, object value2)
            : base(value, value2, ErrorMessage)
        {
        }
    }

    public class RequireBooleanException : ValueConvertException
    {
        static readonly string ErrorMessage = "BOOL 형식(true, false)을 입력해야 합니다.";
        public RequireBooleanException(object value, object value2)
            : base(value, value2, ErrorMessage)
        {
        }
    }
}
