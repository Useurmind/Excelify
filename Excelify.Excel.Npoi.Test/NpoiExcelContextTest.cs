using BaseLibExt.Files;
using Excelify.Core.Types;
using FluentAssertions;
using System;
using Xunit;

namespace Excelify.Excel.Npoi.Test
{
    public class NpoiExcelContextTest
    {
        [Fact]
        public void CreateChangeSaveOpenReadWorks()
        {
            var location = new ExcelCellLocation("sheet1", 1, 3);
            var value = "asdas fsdl.ögkjd";
            var readValue = string.Empty;
            IInMemoryFile excelFile = null;
            
            using (var context1 = new NpoiExcelContext())
            {
                context1.OpenExcel();

                context1.SetValue(location, value);

                excelFile = context1.WriteExcel("myExcel");
            }

            using(var context2 = new NpoiExcelContext())
            {
                context2.OpenExcel(excelFile);

                readValue = context2.GetValue(location);
            }

            readValue.Should().Be(value);
            excelFile.FileName.Should().Be("myExcel.xlsx");
        }
    }
}
