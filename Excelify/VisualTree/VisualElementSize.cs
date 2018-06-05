using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.VisualTree
{
    public struct VisualElementSize
    {
        public VisualElementSize(int rows, int columns)
        {
            this.Rows = rows;
            this.Columns = columns;
        }

        public int Rows { get; private set; }

        public int Columns { get; private set; }

        public override bool Equals(object other)
        {
            if (other != null && other is VisualElementSize)
            {
                var otherIndex = (VisualElementSize)other;
                return this.Rows == otherIndex.Rows && this.Columns == otherIndex.Columns;
            }

            return false;
        }

        public static VisualElementSize operator +(VisualElementSize index1, VisualElementSize index2)
        {
            return new VisualElementSize(index1.Rows + index2.Rows, index1.Columns + index2.Columns);
        }

        public static bool operator ==(VisualElementSize index1, VisualElementSize index2)
        {
            return index1.Equals(index2);
        }

        public static bool operator !=(VisualElementSize index1, VisualElementSize index2)
        {
            return !index1.Equals(index2);
        }
    }
}
