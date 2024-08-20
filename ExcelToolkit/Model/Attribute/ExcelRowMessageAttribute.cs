using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ExcelRowMessageAttribute : Attribute
    {
        public ExcelRowMessageAttribute()
        {
            Title = "错误信息";
        }
        public ExcelRowMessageAttribute(string title = "错误信息", bool ignoreMessage = false)
        {
            Title = title;
            IgnoreMessage = ignoreMessage;
        }

        /// <summary>
        /// 是否忽略错误信息
        /// </summary>
        public bool IgnoreMessage { get; set; }

        /// <summary>
        /// 表头
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 获取提示信息
        /// </summary>
        /// <param name="excelRowInfo"></param>
        /// <returns></returns>
        public virtual string GetMessage(ExcelRowInfo excelRowInfo)
        {
            if (excelRowInfo == null)
                return null;

            if (excelRowInfo.Message != null)
                return excelRowInfo.Message;

            return string.Join(',', excelRowInfo.CellInfos.Where(t => !string.IsNullOrEmpty(t.Message)).Select(t => $"列{t.CellName}:{t.Message}"));
        }
    }
}
