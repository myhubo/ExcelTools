using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Helper
{
    public static partial class ExcelHelper
    {

        public static object GetValue(this ICell cell, Type propertyType)
        {
            try
            {
                if (cell == null)
                    return null;

                var value = cell.CellType == CellType.Blank ? "" : cell.ToString();
                if (propertyType == typeof(long) || propertyType == typeof(long?))
                {
                    if (long.TryParse(value, out long rs))
                        return rs;
                }

                if (propertyType == typeof(decimal) || propertyType == typeof(decimal?))
                {
                    if (decimal.TryParse(value, out decimal rs))
                        return rs;
                }

                if (propertyType == typeof(double) || propertyType == typeof(double?))
                {
                    if (double.TryParse(value, out double rs))
                        return rs;
                }

                if (propertyType == typeof(int) || propertyType == typeof(int?))
                {
                    if (int.TryParse(value, out int rs))
                        return rs;
                }

                if (propertyType == typeof(bool) || propertyType == typeof(bool?))
                {
                    if (bool.TryParse(value, out bool rs))
                        return rs;
                }

                if (propertyType == typeof(DateTime) || propertyType == typeof(DateTime?))
                {
                    if (DateTime.TryParse(value, out DateTime rs))
                        return rs;
                }

                if (propertyType == typeof(string))
                    return value;

                return Convert.ChangeType(value, propertyType);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
