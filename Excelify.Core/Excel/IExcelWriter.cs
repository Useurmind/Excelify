using Excelify.Core.Types;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Core.Excel
{
    public interface IExcelWriter
    {
        void SetValue(ExcelCellLocation cellLocation, string value);
    }
}
