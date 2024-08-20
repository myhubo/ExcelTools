using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    public class ExcelRowData<T> where T : class
    {
        public ExcelRowData(T data, ExcelRowInfo rowInfo)
        {
            Data = data;
            RowInfo = rowInfo;
        }

        /// <summary>
        /// excel行数据
        /// </summary>
        public T Data { get; set; }

        /// <summary>
        /// 行消息
        /// </summary>
        public ExcelRowInfo RowInfo { get; set; }

        public bool HasError => RowInfo != null && RowInfo.HasError;
    }
}