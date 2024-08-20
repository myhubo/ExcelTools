using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{

    /// <summary>
    /// 行信息
    /// </summary>
    public class ExcelRowInfo
    {
        public ExcelRowInfo(int row, string message = "")
        {
            RowNumber = row;
            Message = message;
        }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 列信息
        /// </summary>
        public List<ExcelCellInfo> CellInfos { get; set; } = new();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 是否有错误
        /// </summary>
        public bool HasError => !string.IsNullOrEmpty(Message) || CellInfos.Any(t => !string.IsNullOrEmpty(t.Message));
    }
}
