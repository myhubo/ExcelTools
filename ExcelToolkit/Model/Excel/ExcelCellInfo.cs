using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    /// <summary>
    /// 列信息
    /// </summary>
    public class ExcelCellInfo
    {
        public ExcelCellInfo(int cell, string cellName = "", string cellValue = "", string message = "")
        {
            CellNumber = cell;
            CellName = cellName;
            CellValue = cellValue;
            Message = message;
        }

        /// <summary>
        /// 列号
        /// </summary>
        public int CellNumber { get; set; }

        /// <summary>
        /// 列名
        /// </summary>
        public string CellName { get; set; }

        /// <summary>
        /// 内容
        /// </summary>
        public string CellValue { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string Message { get; set; }
    }
}
