using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    /// <summary>
    /// 声明字段为excel列
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// 按下标读取
        /// </summary>
        /// <param name="index"></param>
        /// <param name="name"></param>
        public ExcelColumnAttribute(int index, string name = "")
        {
            Index = index;
            Name = name;
        }

        /// <summary>
        /// 按名称读取
        /// </summary>
        /// <param name="name"></param>
        /// <param name="index"></param>
        public ExcelColumnAttribute(string name, int index = -1)
        {
            Name = name;
            Index = index;
        }

        /// <summary>
        /// 在excel中的列下标（默认从0开始）
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 在excel中表头名称
        /// </summary>
        public string Name { get; set; }
    }
}
