
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    /// <summary>
    /// 声明实体类为excel模板
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ExcelTemplateAttribute : Attribute
    {
        /// <summary>
        /// 读取excel模板
        /// </summary>
        /// <param name="excelColumnReadType">默认按表头名称读取</param>
        /// <param name="dataStartRow">数据开始的行，默认为1</param>
        /// <param name="headerRow">表头所在行默认为0</param>
        /// <param name="sheetIndex">按sheet下标读取，默认为0</param>
        /// <param name="sheetName">按sheet名称读取，不为空时，sheetIndex字段将无效</param>
        /// <param name="maxColumnIndex">最大列，默认为0表示按实际数据</param>
        public ExcelTemplateAttribute(ExcelColumnReadType excelColumnReadType = ExcelColumnReadType.ColumnName, int dataStartRow = 1, int headerRow = 0, int sheetIndex = 0, string? sheetName = null)
        {
            ExcelColumnReadType = excelColumnReadType;
            DataStartRow = dataStartRow;
            HeaderRow = headerRow;
            SheetIndex = sheetIndex;
            SheetName = sheetName;

            CheckValidity();
        }

        /// <summary>
        /// 检查配置有效性
        /// </summary>
        /// <exception cref="ExcelException"></exception>
        public void CheckValidity()
        {
            if (SheetIndex < 0 && SheetName == null)
                throw new ExcelException("sheetName设置为空时，sheetIndex不能小于0");

            if (ExcelColumnReadType == ExcelColumnReadType.ColumnName)
            {
                if (HeaderRow >= DataStartRow)
                    throw new ExcelException("按字段名读取时，表头所在行必须小于数据开始行");

                if (HeaderRow < 0)
                    throw new ExcelException("按字段名读取时，表头所在行不能小于0");

                if (DataStartRow < 1)
                    throw new ExcelException("按字段名读取时，数据开始行不能小于1");
            }
            else
            {
                if (DataStartRow < 0)
                    throw new ExcelException("按字段名读取时，数据开始行不能小于0");
            }
        }


        /// <summary>
        /// 读取第几个sheet
        /// </summary>
        public int SheetIndex { get; private set; }

        /// <summary>
        /// 读取指定名称的sheet
        /// </summary>
        public string? SheetName { get; private set; }

        /// <summary>
        /// excel列读取类型
        /// </summary>
        public ExcelColumnReadType ExcelColumnReadType { get; private set; }

        /// <summary>
        /// 表头所在行号
        /// </summary>
        public int HeaderRow { get; init; }

        /// <summary>
        /// 数据开始读取行号
        /// </summary>
        public int DataStartRow { get; private set; }
    }

    /// <summary>
    /// excel列读取类型
    /// </summary>
    public enum ExcelColumnReadType
    {
        /// <summary>
        /// 按列名读取
        /// </summary>
        ColumnName = 0,

        /// <summary>
        /// 按列下标读取
        /// </summary>
        ColumnIndex = 1,
    }
}
