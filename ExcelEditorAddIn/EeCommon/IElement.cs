using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeCommon
{
    public enum ElementType
    {
        Value,
        Array,
        Object,
        /// <summary> Array containing objects </summary>
        Table,
    }

    public interface IElement
    {
        ElementType Type { get; }
        object GetValue();
        string GetJsonText();
    }

    public interface IValueElement
    {
    }

    public interface IArrayElement : IElement
    {
        int Length { get; }
        bool Any { get; }
        bool Empty { get; }
    }

    public interface IObjectElement : IElement
    {
        List<string> Keys { get; }
    }

    public interface ITableElement : IElement, IArrayElement
    {
        List<string> Keys { get; }
        object[,] Values { get; }
        IElement[,] Elements { get; }
    }
}
