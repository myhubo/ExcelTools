using ExcelToolkit.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject1
{
    /// <summary>
    /// 测试数据2
    /// 按下标读取，读取名为“sheet2”的数据
    /// </summary>
    [ExcelTemplate(ExcelColumnReadType.ColumnIndex,sheetName:"sheet2")]
    public class TestData2
    {
        /// <summary>
        /// 编码
        /// </summary>
        [ExcelColumn(0,"编码")]
        public string Code { get; set; }

        /// <summary>
        /// 商品名称
        /// </summary>
        [ExcelColumn(1,"商品名称")]
        public string Name { get; set; }

        /// <summary>
        /// 型号
        /// </summary>
        [ExcelColumn(2,"型号")]
        public string Models { get; set; }

        /// <summary>
        /// 价格
        /// </summary>
        [ExcelColumn(3,"价格")]
        public Decimal Price { get; set; }

        /// <summary>
        /// 数量
        /// </summary>
        [ExcelColumn(4,"数量")]
        public int Number { get; set; }

        /// <summary>
        /// 单位
        /// </summary>
        [ExcelColumn(5,"单位")]
        public string UnitName { get; set; }
    }
}
