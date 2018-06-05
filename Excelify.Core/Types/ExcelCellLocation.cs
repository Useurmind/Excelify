using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Core.Types
{
    /// <summary>
    /// Absolute location of a cell in an excel workbook.
    /// Sheet + cell index.
    /// </summary>
    public struct ExcelCellLocation
    {
        public ExcelCellLocation(string sheetName, int row, int column): this(sheetName, new ExcelCellIndex(row, column))
        {
        }

        public ExcelCellLocation(string sheetName, ExcelCellIndex cellIndex)
        {
            this.SheetName = sheetName;
            this.CellIndex = cellIndex;
        }

        public string SheetName { get; private set; }

        public ExcelCellIndex CellIndex { get; private set; }

        public override bool Equals(object other)
        {
            if (other != null && other is ExcelCellLocation)
            {
                var otherIndex = (ExcelCellLocation)other;
                return this.SheetName == otherIndex.SheetName && this.CellIndex == otherIndex.CellIndex;
            }

            return false;
        }

        public static ExcelCellLocation operator +(ExcelCellLocation loc, ExcelCellIndex index)
        {
            return new ExcelCellLocation(loc.SheetName, loc.CellIndex + index);
        }

        public static bool operator==(ExcelCellLocation loc1, ExcelCellLocation loc2)
        {
            return loc1.Equals(loc2);
        }

        public static bool operator!=(ExcelCellLocation loc1, ExcelCellLocation loc2)
        {
            return !loc1.Equals(loc2);
        }
    }
}
