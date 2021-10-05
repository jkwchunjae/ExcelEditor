using EeCommon;
using System.Collections.Generic;

namespace ExcelEditorAddIn
{
    public class ColumnComparer : IComparer<(string PropertyName, ElementType ElementType)>
    {
        ColumnSetting _columnSetting;
        public ColumnComparer(ColumnSetting columnSetting)
        {
            _columnSetting = columnSetting;
        }

        public int Compare((string PropertyName, ElementType ElementType) a, (string PropertyName, ElementType ElementType) b)
        {
            var aSetting = _columnSetting?.GetOrderData(a.PropertyName);
            var bSetting = _columnSetting?.GetOrderData(b.PropertyName);

            return (aSetting?.Order ?? int.MaxValue) - (bSetting?.Order ?? int.MaxValue);
        }
    }
}
