using Excelify.Core.Types;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Core.Excel
{
    public interface IExcelReader
    {
        string GetValue(ExcelCellLocation cellLocation);
    }
}
