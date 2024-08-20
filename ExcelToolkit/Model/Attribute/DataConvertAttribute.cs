using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    /// <summary>
    /// 数据转换
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public abstract class DataConvertAttribute : Attribute
    {

        public abstract object ConvertValue(object value);
    }
}
