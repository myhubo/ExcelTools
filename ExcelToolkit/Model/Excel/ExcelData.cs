using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    public class ExcelData<T> where T : class
    {
        /// <summary>
        /// excel中的数据
        /// </summary>
        private List<ExcelRowData<T>> _data;
        public ExcelData(int count)
        {
            _data = new List<ExcelRowData<T>>(count);
        }

        /// <summary>
        /// 按下标取数据
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelRowData<T> this[int index] => _data[index];

        /// <summary>
        /// 添加数据
        /// </summary>
        /// <param name="rowData"></param>
        public void Add(ExcelRowData<T> rowData) => _data.Add(rowData);

        /// <summary>
        /// 清除数据
        /// </summary>
        public void Clear() => _data.Clear();

        /// <summary>
        /// 获取数据列表
        /// </summary>
        public IEnumerable<T> Data => _data.Select(t => t.Data);

        /// <summary>
        /// 获取带状态数据列表
        /// </summary>
        public IEnumerable<ExcelRowData<T>> ExcelRowData => _data;

        /// <summary>
        /// 是否有错误
        /// </summary>
        public bool HasError => _data.Any(t => t.HasError);

        /// <summary>
        /// 数量
        /// </summary>
        public int Count => _data.Count;

        /// <summary>
        /// 错误数据条数
        /// </summary>
        public int ErrorCount => _data.Count(t => t.HasError);

        /// <summary>
        /// 获取错误数据
        /// </summary>
        public IEnumerable<ExcelRowData<T>> ErrorExcelRowData => _data.Where(t => t.HasError);
    }
}
