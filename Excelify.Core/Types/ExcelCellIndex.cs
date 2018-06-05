using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Core.Types
{
    /// <summary>
    /// A simple combination of row and column.
    /// </summary>
    public struct ExcelCellIndex
    {
        public ExcelCellIndex(int row, int column)
        {
            this.Row = row;
            this.Column = column;
        }

        public int Row { get; private set; }

        public int Column { get; private set; }

        public override bool Equals(object other)
        {
            if (other != null && other is ExcelCellIndex)
            {
                var otherIndex = (ExcelCellIndex)other;
                return this.Row == otherIndex.Row && this.Column == otherIndex.Column;
            }

            return false;
        }

        public static ExcelCellIndex operator+(ExcelCellIndex index1, ExcelCellIndex index2)
        {
            return new ExcelCellIndex(index1.Row + index2.Row, index1.Column + index2.Column);
        }

        public static bool operator==(ExcelCellIndex index1, ExcelCellIndex index2)
        {
            return index1.Equals(index2);
        }

        public static bool operator!=(ExcelCellIndex index1, ExcelCellIndex index2)
        {
            return !index1.Equals(index2);
        }
    }
}
