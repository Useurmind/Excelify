using BaseLibExt.Files;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Core.Excel
{
    public interface IExcelStorage
    {
        void OpenExcel(IInMemoryFile excelFile = null);
        
        IInMemoryFile WriteExcel(string fileName = null, string contentType = null);
    }
}
