using ExcelToolkit.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestProject1.Dto;

namespace TestProject1
{
    /// <summary>
    /// 测试数据1
    /// 按字段名读取，读取第0个sheet
    /// </summary>
    [ExcelTemplate]
    [MyExcelRowMessage]
    public class TestData1
    {
        [Required]
        /// <summary>
        /// 编码
        /// </summary>
        [ExcelColumn("编码")]
        public string Code { get; set; }

        /// <summary>
        /// 商品名称
        /// </summary>
        [ExcelColumn("商品名称")]
        public string Name { get; set; }

        /// <summary>
        /// 型号
        /// </summary>
        [ExcelColumn("型号")]
        public string Models { get; set; }

        /// <summary>
        /// 价格
        /// </summary>
        [ExcelColumn("价格")]
        [DecimalDataConvert]
        public decimal Price { get; set; }

        /// <summary>
        /// 数量
        /// </summary>
        [ExcelColumn("数量")]
        public int Number { get; set; }

        /// <summary>
        /// 单位
        /// </summary>
        [ExcelColumn("单位")]
        public string UnitName { get; set; }
    }
}
