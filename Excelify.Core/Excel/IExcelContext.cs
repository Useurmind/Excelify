using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Core.Excel
{
    public interface IExcelContext : IExcelReader, IExcelWriter, IExcelStorage
    {
    }
}
