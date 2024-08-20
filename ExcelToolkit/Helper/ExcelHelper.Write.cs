using ExcelToolkit.Model;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Helper
{
    public static partial class ExcelHelper
    {
        /// <summary>
        /// 导出数据
        /// </summary>
        /// <typeparam name="T">输出的数据类型</typeparam>
        /// <param name="data">数据集合</param>
        /// <param name="isXlsx">是否为2007格式，否则为2003格式.xls</param>
        /// <param name="sheetName">输出的sheet名称</param>
        /// <returns></returns>
        /// <exception cref="ExcelException"></exception>
        public static byte[] Export<T>(this IEnumerable<T> data, bool isXlsx = true, string sheetName = "Sheet1") where T : class
        {
            var properties = typeof(T).GetProperties().Where(t => t.CustomAttributes.Any(a => a.AttributeType == typeof(ExcelColumnAttribute)));
            var propertyDict = new Dictionary<int, PropertyInfo>();
            var titleDict = new Dictionary<int, string>();// 标题缓存

            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<ExcelColumnAttribute>();
                var index = attribute.Index;
                if (index == -1)
                    throw new ExcelException($"文件模板列{attribute.Name}序号未配置");

                if (propertyDict.ContainsKey(index))
                    throw new ExcelException("文件模板列{attribute.Name}序号配置重复");

                propertyDict.Add(index, property);
                titleDict.Add(index, attribute.Name ?? property.Name);
            }

            IWorkbook book = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook(); // XSSFWorkbook 2007格式 HSSFWorkbook 2003格式
            var sheet = book.CreateSheet(sheetName);
            properties = propertyDict.OrderBy(t => t.Key).Select(t => t.Value).ToList();

            var rowIndex = 0;
            var cellIndex = 0;

            #region 设置表头

            var headerRow = sheet.CreateRow(rowIndex++);
            foreach (var item in titleDict.OrderBy(t => t.Key))
            {
                var cell = headerRow.CreateCell(cellIndex++);
                cell.SetCellValue(item.Value);
            }

            #endregion

            foreach (var t in data)
            {
                var row = sheet.CreateRow(rowIndex++);
                var index = 0;
                foreach (var item in propertyDict.OrderBy(t => t.Key))
                {
                    var property = item.Value;
                    var cell = row.CreateCell(index++);
                    cell.SetCellValue(property.GetValue(t, null)?.ToString());
                }
            }

            using (var ms = new ExcelMemoryStream())
            {
                book.Write(ms);
                ms.Seek(0, SeekOrigin.Begin);
                ms.Dispose();

                return ms.ToArray();
            }
        }
    }
}
