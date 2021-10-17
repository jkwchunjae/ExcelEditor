using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeCommon
{
    public enum ElementType
    {
        Null,
        Value,
        Array,
        Object,
        /// <summary> Array containing objects </summary>
        Table,
    }

    public enum ValueType
    {
        Null,
        String,
        Integer,
        Boolean,
        DateTime,
        Float,
    }

    public interface IHelper
    {
        T Deserialize<T>(string text);
        string Serialize<T>(T obj);
        string MetadataFilePath(string filePath);
    }

    public interface IElement : IHelper
    {
        ElementType Type { get; }
        object GetExcelValue();
        string GetSaveText();
        IValueElement CreateValueElement(object value, object value2);
        IValueElement CreateValueElement(object value, object value2, ValueType valueType);
    }

    public interface IValueElement : IElement
    {
        ValueType ValueType { get; }
        void UpdateValue(object value, object value2);
    }

    public interface IArrayElement : IElement
    {
        int Length { get; }
        bool Any { get; }
        bool Empty { get; }

        IEnumerable<IValueElement> Elements { get; }

        void Add(IElement element);
        void AddAt(int index, IElement element);
        void RemoveAt(int index);
    }

    public interface IObjectElement : IElement
    {
        IReadOnlyDictionary<string, IElement> Properties { get; }
        void Add(string key, IElement value);
        void Remove(string key);
    }

    public interface ITableElement : IElement
    {
        int Length { get; }
        bool Any { get; }
        bool Empty { get; }
        IEnumerable<IObjectElement> Elements { get; }
        IEnumerable<(string PropertyName, ElementType ElementType)> Properties { get; }

        void Add(IObjectElement objectElement);
        void AddAt(int index, IObjectElement objectElement);
        void RemoveAt(int index);
    }
}
