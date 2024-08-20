using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Model
{
    public class ExcelMemoryStream : MemoryStream
    {
        private bool _allowClose { get; set; }
        public ExcelMemoryStream(bool allowClose = false)
        {
            _allowClose = allowClose;
        }

        public override void Close()
        {
            if (_allowClose)
                base.Close();
        }

        public new void Dispose()
        {
            _allowClose = true;
        }
    }
}
