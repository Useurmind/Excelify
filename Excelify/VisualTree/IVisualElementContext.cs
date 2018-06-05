using Excelify.Core.Types;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.VisualTree
{
    public interface IVisualElementContext
    {
        object Model { get; }

        ExcelCellLocation RootLocation { get; }
    }
}
