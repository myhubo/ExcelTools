# ExcelTools
基于NPOI快捷读取和输出excel的小工具

1. 支持读取.xlsx和.xls格式的excel文件为对象列表
2. 支持将对象列表输出为.xlsx和.xls格式的excel文件
3. 如果读取的excel有错误，可在原文件最后增加具体错误信息

# 定义实体类
- 定义实体对象，通过 `ExcelTemplateAttribute` 指定读取方式，按标题读取或按位置读取；
- 通过 `ExcelRowMessageAttribute` 自定义错误提示格式

```csharp
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
    
        /// …… 省略其他属性
    }
```

# 读取Excel
- 读取Excel文件，并返回对象列表
```csharp 
   var data = ExcelHelper.ReadData<TestData1>(path);
```
- 读取Excel文件，在原excel文件中加入错误信息列
```csharp 
   (_,var stream) = ExcelHelper.ReadDataAndExport<TestData1>(path);
```
# 生成excel
```csharp 
    var data = new List<TestData1>();
    data.Add(new TestData1() { Code = "123", Name = "测试" });
    data.Add(new TestData1() { Code = "123", Name = "测试" });
    var bytes = data.Export();
```

# 其他配置
- 自定义数据转换`DataConvertAttribute`
```csharp
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

[ExcelTemplate]
[MyExcelRowMessage]
public class TestData1
{
    /// <summary>
    /// 价格
    /// </summary>
    [ExcelColumn("价格")]
    [DecimalDataConvert]
    public decimal Price { get; set; }
}
```

# 具体用法参照TestProject
