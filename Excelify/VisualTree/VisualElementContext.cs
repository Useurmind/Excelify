using Excelify.Core.Types;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.VisualTree
{
    public class VisualElementContext : IVisualElementContext
    {
        public VisualElementContext(object model, ExcelCellLocation rootLocation)
        {
            this.Model = model;
            this.RootLocation = rootLocation;
        }

        public object Model { get; }
        public ExcelCellLocation RootLocation { get; }
    }
}
