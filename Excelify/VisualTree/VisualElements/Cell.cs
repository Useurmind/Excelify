using System;
using System.Collections.Generic;
using System.Text;
using Excelify.Core.Excel;

namespace Excelify.VisualTree.VisualElements
{
    /// <summary>
    /// A simple excel cell.
    /// </summary>
    public class Cell : IVisualElement
    {
        public string Path { get; private set; }

        public Cell(string path)
        {
            this.Path = path;
        }

        public VisualElementSize GetSize(object model)
        {
            return new VisualElementSize(1, 1);
        }

        public void Read(IExcelReader reader, IVisualElementContext context)
        {
            var stringValue = reader.GetValue(context.RootLocation);

            if(context.Model != null)
            {

            }
        }

        public void Write(IExcelWriter writer, IVisualElementContext context)
        {
            throw new NotImplementedException();
        }
    }
}
