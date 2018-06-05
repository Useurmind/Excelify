using Excelify.Core.Types;
using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Excelify.Core.Test.Types
{
    public class ExcelCellIndexTest
    {
        [Theory]
        [InlineData(1, 2, 1, 2, true)]
        [InlineData(5, 9, 5, 9, true)]
        [InlineData(7, 3, 3, 7, false)]
        [InlineData(1, 2, 3, 4, false)]
        public void EqualsWorksForOtherCellIndexes(int row1, int column1, int row2, int column2, bool equals)
        {
            var index1 = new ExcelCellIndex(row1, column1);
            var index2 = new ExcelCellIndex(row2, column2);

            var result = index1.Equals(index2);

            result.Should().Be(equals);
        }

        [Fact]
        public void EqualsReturnsFalseForNonIndex()
        {
            var index = new ExcelCellIndex(1, 2);
            
            var result = index.Equals("asdas");

            result.Should().Be(false);
        }
    }
}
