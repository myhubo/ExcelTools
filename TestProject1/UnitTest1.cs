using ExcelToolkit.Helper;
using System.IO;

namespace TestProject1
{
    [TestClass]
    public class UnitTest1
    {
        /// <summary>
        /// 测试读取xlsx，按字段名读取
        /// </summary>
        [TestMethod]
        public void TestReadXlsxWithColumnName()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xlsx");

            var data = ExcelHelper.ReadData<TestData1>(path);
      
            Assert.AreEqual(data.Count, 830);      // 总条数
            Assert.AreEqual(data.ErrorCount, 3);      // 错误条数
        }

        /// <summary>
        /// 测试读取xlsx，按位置读取
        /// </summary>
        [TestMethod]
        public void TestReadXlsxWithColumnIndex()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xlsx");

            var data = ExcelHelper.ReadData<TestData2>(path);

            Assert.AreEqual(data.Count, 830);      // 总条数
            Assert.AreEqual(data.ErrorCount, 4);      // 错误条数
        }


        /// <summary>
        /// 测试导出xlsx
        /// </summary>
        [TestMethod]
        public void TestWriteXlsx()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xlsx");

            var data = ExcelHelper.ReadData<TestData2>(path);

            var bytes = data.Data.Export();
            var filepath = Path.Combine(exeDirectory, "Tests//导出_测试数据1.xlsx");
            string directoryPath = Path.GetDirectoryName(filepath);
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            // 将字节数组写入到指定的文件路径
            File.WriteAllBytes(filepath, bytes);
        }

        /// <summary>
        /// 测试读取并导出xlsx
        /// </summary>
        [TestMethod]
        public void TestReadAndExportXlsx()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xlsx");

            (_,var stream) = ExcelHelper.ReadDataAndExport<TestData1>(path);

            var filepath = Path.Combine(exeDirectory, "Tests//导出_测试数据(带错误提示).xlsx");
            string directoryPath = Path.GetDirectoryName(filepath);
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            // 将字节数组写入到指定的文件路径
            using (var fileStream = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                stream.CopyTo(fileStream);

            stream.Dispose();
        }


        /// <summary>
        /// 测试读取xls，按字段名读取
        /// </summary>
        [TestMethod]
        public void TestReadXlsWithColumnName()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xls");

            var data = ExcelHelper.ReadData<TestData1>(path);

            Assert.AreEqual(data.Count, 830);      // 总条数
            Assert.AreEqual(data.ErrorCount, 3);      // 错误条数
        }

        /// <summary>
        /// 测试读取xls，按位置读取
        /// </summary>
        [TestMethod]
        public void TestReadXlsWithColumnIndex()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xls");

            var data = ExcelHelper.ReadData<TestData2>(path);

            Assert.AreEqual(data.Count, 830);      // 总条数
            Assert.AreEqual(data.ErrorCount, 4);      // 错误条数
        }


        /// <summary>
        /// 测试导出xlsx
        /// </summary>
        [TestMethod]
        public void TestWriteXls()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xls");

            var data = ExcelHelper.ReadData<TestData2>(path);

            var bytes = data.Data.Export();
            var filepath = Path.Combine(exeDirectory, "Tests//导出_测试数据1.xls");
            string directoryPath = Path.GetDirectoryName(filepath);
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            // 将字节数组写入到指定的文件路径
            File.WriteAllBytes(filepath, bytes);
        }

        /// <summary>
        /// 测试读取并导出xlsx
        /// </summary>
        [TestMethod]
        public void TestReadAndExportXls()
        {
            var exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var path = Path.Combine(exeDirectory, "Tests//测试数据1.xls");

            (_, var stream) = ExcelHelper.ReadDataAndExport<TestData1>(path);

            var filepath = Path.Combine(exeDirectory, "Tests//导出_测试数据(带错误提示).xls");
            string directoryPath = Path.GetDirectoryName(filepath);
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            // 将字节数组写入到指定的文件路径
            using (var fileStream = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                stream.CopyTo(fileStream);

            stream.Dispose();
        }
    }
}