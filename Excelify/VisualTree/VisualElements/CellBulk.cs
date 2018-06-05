using System;
using System.Collections.Generic;
using System.Text;
using Excelify.Core.Excel;

namespace Excelify.VisualTree.VisualElements
{
    /// <summary>
    /// A bulk of cells that can be formed up in an arbitrary manner (reactangle, sparse matrix, etc.)
    /// </summary>
    public class CellBulk : IVisualElement
    {
        public string Path => throw new NotImplementedException();

        public VisualElementSize GetSize(object model)
        {
            throw new NotImplementedException();
        }

        public void Read(IExcelReader reader, IVisualElementContext context)
        {
            throw new NotImplementedException();
        }

        public void Write(IExcelWriter writer, IVisualElementContext context)
        {
            throw new NotImplementedException();
        }
    }
}
