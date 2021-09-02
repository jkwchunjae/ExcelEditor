using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeCommon
{
    public enum DocumentType
    {
        Value,
        Array,
        Object,
        /// <summary> Array containing objects </summary>
        Table,
    }

    public interface IDocument
    {
        DocumentType Type { get; }
        string GetString();
    }

    public interface IArrayDocument : IDocument
    {
        int Length { get; }
        bool Any { get; }
        bool Empty { get; }
    }

    public interface IObjectDocument : IDocument
    {
        List<string> Keys { get; }
    }

    public interface ITableDocument : IDocument, IArrayDocument
    {
        List<string> Keys { get; }
        object[,] Values { get; }
    }
}
