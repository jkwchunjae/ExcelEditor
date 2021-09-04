using JkwExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeCommon
{
    public static class ITableElementExtensions
    {
        /// <summary>
        /// table 데이터를 엑셀 시트에 표시하도록 2차원 배열로 리턴한다.
        /// </summary>
        /// <param name="tableElement"></param>
        /// <param name="propertyNameOrder">
        ///     property 순서를 정해준다. 엑셀 시트의 순서가 객체의 순서와 다를 수 있다.
        ///     없는게 있을 수도 있고.
        /// </param>
        /// <returns></returns>
        public static object[,] GetExcelArray(
            this ITableElement tableElement, IEnumerable<string> propertyNameOrder)
        {
            var columnDic = propertyNameOrder
                .Select((propertyName, index) => new { propertyName, index })
                .ToDictionary(x => x.propertyName, x => x.index);

            var result = new object[tableElement.Length, propertyNameOrder.Count()];

            tableElement.Elements.ForEach((objectElement, index) =>
            {
                var row = index;
                objectElement.Properties.ForEach(property =>
                {
                    if (columnDic.TryGetValue(property.Key, out var column))
                    {
                        result[row, column] = property.Value.GetExcelValue();
                    }
                });
            });

            return result;
        }
    }
}
