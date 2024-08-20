using ExcelToolkit.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace TestProject1.Dto
{
    public class MyExcelRowMessageAttribute : ExcelRowMessageAttribute
    {
        public override string GetMessage(ExcelRowInfo excelRowInfo)
        {
            return $"读取错误:{JsonConvert.SerializeObject(excelRowInfo)}";
        }
    }
}
