using BaseLibExt.Files;
using Excelify.Core.Excel;
using Excelify.Core.Types;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Excelify.Excel.Npoi
{
    public sealed class NpoiExcelContext : IExcelContext, IDisposable
    {
        private XSSFWorkbook workbook;

        public void Dispose()
        {
            if (this.workbook != null)
            {
                this.workbook.Close();
            }
        }

        public string GetValue(ExcelCellLocation cellLocation)
        {
            this.CheckWorkbook();

            ISheet sheet = this.workbook.GetSheet(cellLocation.SheetName);
            if (sheet == null)
            {
                return string.Empty;
            }

            IRow row = sheet.GetRow(cellLocation.CellIndex.Row);
            if (row == null)
            {
                return string.Empty;
            }

            var cell = row.GetCell(cellLocation.CellIndex.Column, MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell == null)
            {
                return string.Empty;
            }

            return cell.StringCellValue;
        }

        public void OpenExcel(IInMemoryFile excelFile = null)
        {
            if (excelFile != null)
            {
                this.workbook = new XSSFWorkbook(excelFile.DataStream);
            }
            else
            {
                this.workbook = new XSSFWorkbook();
            }
        }

        public IInMemoryFile WriteExcel(string fileName, string contentType = null)
        {
            this.CheckWorkbook();

            var memoryStream = new MemoryStream();
            this.workbook.Write(memoryStream);
            var memoryStream2 = new MemoryStream(memoryStream.ToArray());
            return InMemoryFileFactory.CreateFromStream(memoryStream2, $"{fileName}.xlsx", contentType);
        }

        public void SetValue(ExcelCellLocation cellLocation, string value)
        {
            this.CheckWorkbook();

            var cell = GetCellEnsured(cellLocation);

            cell.SetCellValue(value);
        }

        private void CheckWorkbook()
        {
            if (this.workbook == null)
            {
                throw new Exception("No workbook open in NpoiExcelContext");
            }
        }

        private ISheet GetSheetEnsured(string sheetName)
        {
            this.CheckWorkbook();

            var sheet = this.workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = this.workbook.CreateSheet(sheetName);
            }

            return sheet;
        }

        private IRow GetRowEnsured(string sheetName, int rowIndex)
        {
            var sheet = this.GetSheetEnsured(sheetName);

            IRow row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }

            return row;
        }

        private ICell GetCellEnsured(ExcelCellLocation cellLocation)
        {
            return this.GetCellEnsured(cellLocation.SheetName, cellLocation.CellIndex.Row, cellLocation.CellIndex.Column);
        }

        private ICell GetCellEnsured(string sheetName, int rowIndex, int columnIndex)
        {
            var row = this.GetRowEnsured(sheetName, rowIndex);

            ICell cell = row.GetCell(columnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell == null)
            {
                cell = row.CreateCell(columnIndex, CellType.String);
            }

            return cell;
        }
    }
}
