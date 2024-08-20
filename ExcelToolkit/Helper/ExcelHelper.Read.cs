using ExcelToolkit.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.VisualBasic;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util.Collections;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Helper
{
    public static partial class ExcelHelper
    {
        public static ExcelData<T> ReadData<T>(string filepath, Func<T, string> validateRowFunc = null) where T : class
        {
            var isXlsx = filepath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
            using (var stream = File.OpenRead(filepath))
                return ReadData(stream, isXlsx, validateRowFunc).data;
        }

        public static ExcelData<T> ReadData<T>(IFormFile file, Func<T, string> validateRowFunc = null) where T : class
        {
            var isXlsx = file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
            using (var stream = file.OpenReadStream())
                return ReadData(stream, isXlsx, validateRowFunc).data;
        }

        public static (ExcelData<T>, ExcelMemoryStream?) ReadDataAndExport<T>(string filepath, Func<T, string> validateRowFunc = null) where T : class
        {
            var isXlsx = filepath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
            using (var stream = File.OpenRead(filepath))
                return ReadData(stream, isXlsx, validateRowFunc, true);
        }

        public static (ExcelData<T>, ExcelMemoryStream?) ReadDataAndExport<T>(IFormFile file, Func<T, string> validateRowFunc = null) where T : class
        {
            var isXlsx = file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
            using (var stream = file.OpenReadStream())
                return ReadData(stream, isXlsx, validateRowFunc, true);
        }

        public static (ExcelData<T> data, ExcelMemoryStream? stream) ReadData<T>(Stream stream, bool isXlsx = true, Func<T, string> validateRowFunc = null, bool exportIfHasError = false) where T : class
        {
            try
            {
                var excelTemplateAttribute = typeof(T).GetCustomAttribute<ExcelTemplateAttribute>();
                if (excelTemplateAttribute == null)
                    throw new ExcelException($"类型{typeof(T)}必须实现ExcelTemplateAttribute");

                excelTemplateAttribute.CheckValidity();

                #region 读取列对应属性

                var messageAttribute = typeof(T).GetCustomAttribute<ExcelRowMessageAttribute>() ?? new ExcelRowMessageAttribute();
                var properties = typeof(T).GetProperties().Where(t => t.CustomAttributes.Any(a => a.AttributeType == typeof(ExcelColumnAttribute)));
                var propertyInfoDict = new Dictionary<string, PropertyInfo>(); // cell绑定属性缓存
                var dataConvertDict = new Dictionary<string, DataConvertAttribute>(); // 属性绑定数据转换缓存
                var validateDict = new Dictionary<string, IEnumerable<ValidationAttribute>>(); // 属性绑定数据校验缓存
                var titleDict = new Dictionary<int, string>(); // 标题缓存

                foreach (var property in properties)
                {
                    var excelColumnAttribute = property.GetCustomAttribute<ExcelColumnAttribute>();
                    var key = excelTemplateAttribute.ExcelColumnReadType == ExcelColumnReadType.ColumnIndex ? excelColumnAttribute.Index.ToString() : excelColumnAttribute.Name;
                    if (propertyInfoDict.ContainsKey(key))
                        throw new ExcelException($"列{property.Name}序号{key}配置重复");

                    propertyInfoDict[key] = property;

                    // 按下标读取，缓存标题
                    if (excelTemplateAttribute.ExcelColumnReadType == ExcelColumnReadType.ColumnIndex)
                        titleDict[excelColumnAttribute.Index] = excelColumnAttribute.Name ?? property.Name;

                    // 初始化属性绑定数据转换缓存
                    var convertDataAttribute = property.GetCustomAttribute<DataConvertAttribute>();
                    if (convertDataAttribute != null)
                        dataConvertDict.Add(property.Name, convertDataAttribute);

                    // 初始化属性绑定数据校验缓存
                    var validateAttributes = property.GetCustomAttributes<ValidationAttribute>();
                    if (validateAttributes.Any())
                        validateDict.Add(key, validateAttributes);
                }

                #endregion

                IWorkbook workbook = isXlsx ? new XSSFWorkbook(stream) : new HSSFWorkbook(stream);
                var sheetIndex = excelTemplateAttribute.SheetIndex;
                if (!string.IsNullOrEmpty(excelTemplateAttribute.SheetName))
                    sheetIndex = workbook.GetSheetIndex(excelTemplateAttribute.SheetName);
                var sheet = workbook.GetSheetAt(sheetIndex);
                var lastCell = 0;// 最后一列

                #region 缓存表头名称-按列名称读取

                if (excelTemplateAttribute.ExcelColumnReadType == ExcelColumnReadType.ColumnName)
                {
                    var row = sheet.GetRow(excelTemplateAttribute.HeaderRow);
                    for (var i = row.FirstCellNum; i < row.LastCellNum; i++)
                        titleDict[i] = row.GetCell(i).StringCellValue;
                    lastCell = row.LastCellNum;

                    if (exportIfHasError)
                    {
                        // 设置提示信息表头
                        var messageTitleCell = row.CreateCell(lastCell);
                        messageTitleCell.SetCellValue(messageAttribute.Title);
                    }
                }

                #endregion

                var result = new ExcelData<T>(sheet.LastRowNum - excelTemplateAttribute.DataStartRow + 1);
                for (var i = excelTemplateAttribute.DataStartRow; i <= sheet.LastRowNum; i++)
                {
                    var rowInfo = new ExcelRowInfo(i);
                    T item = null;
                    IRow row = null;
                    try
                    {
                        item = (T)Activator.CreateInstance(typeof(T));
                        row = sheet.GetRow(i);
                        if (row != null)
                        {
                            lastCell = Math.Max(lastCell, row.LastCellNum);
                            for (var j = 0; j < lastCell; j++)
                            {
                                if (!titleDict.ContainsKey(j))
                                    continue;

                                var cellInfo = new ExcelCellInfo(j, titleDict[j]);
                                try
                                {
                                    var cell = row.GetCell(j);

                                    #region 从缓存中获取列对应属性

                                    var propertyInfoKey = "";
                                    PropertyInfo property = null;
                                    if (excelTemplateAttribute.ExcelColumnReadType == ExcelColumnReadType.ColumnIndex)
                                        propertyInfoKey = j.ToString();
                                    else if (excelTemplateAttribute.ExcelColumnReadType == ExcelColumnReadType.ColumnName && titleDict.ContainsKey(j))
                                        propertyInfoKey = titleDict[j];

                                    if (propertyInfoDict.ContainsKey(propertyInfoKey))
                                        property = propertyInfoDict[propertyInfoKey];

                                    #endregion

                                    #region 校验数据格式

                                    if (validateDict.ContainsKey(propertyInfoKey))
                                    {
                                        foreach (var validate in validateDict[propertyInfoKey])
                                        {
                                            if (!validate.IsValid(cell))
                                            {
                                                cellInfo.Message = $"数据校验错误：{validate.ErrorMessage ?? validate.FormatErrorMessage(titleDict[j])}";
                                                rowInfo.CellInfos.Add(cellInfo);
                                                continue;
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 读取列绑定的值

                                    if (property != null)
                                    {
                                        object value = null;
                                        if (dataConvertDict.ContainsKey(property.Name))
                                        {
                                            try
                                            {
                                                value = dataConvertDict[property.Name].ConvertValue(cell);
                                            }
                                            catch (ExcelException ex)
                                            {
                                                cellInfo.Message = $"自定义数据转换出错:{ex.CustomerMessage}";
                                                rowInfo.CellInfos.Add(cellInfo);
                                                continue;
                                            }
                                            catch (Exception ex)
                                            {
                                                cellInfo.Message = $"自定义数据转换出错:{ex.Message}";
                                                rowInfo.CellInfos.Add(cellInfo);
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            try
                                            {
                                                value = cell.GetValue(property.PropertyType);
                                            }
                                            catch (Exception ex)
                                            {
                                                cellInfo.Message = $"读取数据出错:{ex.Message}";
                                                rowInfo.CellInfos.Add(cellInfo);
                                                continue;
                                            }
                                        }
                                        property.SetValue(item, value);
                                    }

                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    rowInfo.Message = "处理列失败：" + ex.Message;
                                }
                            }

                            // 执行自定义的数据校验
                            if (validateRowFunc != null)
                                rowInfo.Message = validateRowFunc(item);
                        }
                    }
                    catch (Exception ex)
                    {
                        rowInfo.Message = $"处理行数据失败:{ex.Message}";
                    }

                    if (!rowInfo.HasError)
                    {
                        rowInfo = null;
                    }
                    else if (exportIfHasError && row != null)
                    {
                        // 设置错误信息
                        var messageDataCell = row.CreateCell(lastCell);
                        messageDataCell.SetCellValue(messageAttribute.GetMessage(rowInfo));
                    }

                    result.Add(new ExcelRowData<T>(item, rowInfo));
                }

                if (exportIfHasError)
                {
                    var outStream = new ExcelMemoryStream();
                    workbook.Write(outStream);
                    outStream.Position = 0; // 重置流位置，便于后续读取
                    return (result, outStream);
                }

                return (result, null);
            }
            catch (ExcelException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                throw new ExcelException("读取文件失败");
            }
        }
    }
}
