using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    /// <summary>
    /// decimal类型转换
    /// </summary>
    public class DecimalDataConvertAttribute : DataConvertAttribute
    {
        public override object ConvertValue(object value)
        {
            if (value == null)
                return value;

            return decimal.Parse(value.ToString());
        }
    }
}
