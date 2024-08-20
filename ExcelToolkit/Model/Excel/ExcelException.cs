using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    public class ExcelException : Exception
    {
        public string CustomerMessage { get; set; }
        public string Description { get; set; }

        public ExcelException(string message = null, string description = null)
        {
            CustomerMessage = message;
            Description = description;
        }
    }
}
