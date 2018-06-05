using Excelify.Core.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.VisualTree
{
    /// <summary>
    /// The core of the visual tree that will be bound to the excel sheet.
    /// </summary>
    public interface IVisualElement
    {
        /// <summary>
        /// The path to the property of the model that is bound to the element.
        /// Null/empty if bound to the model itself.
        /// </summary>
        string Path { get; }

        /// <summary>
        /// Get the size of the visual element in rows and columns of excel sheet cells.
        /// </summary>
        /// <param name="model">The model based on which the size is computed.</param>
        /// <returns>The size of itself.</returns>
        VisualElementSize GetSize(object model);

        /// <summary>
        /// Write this visual element to the given writer.
        /// </summary>
        /// <param name="writer">The writer to write to.</param>
        /// <param name="context">Context information required for the rendering process.</param>
        void Write(IExcelWriter writer, IVisualElementContext context);

        /// <summary>
        /// Read this visual element from the given reader.
        /// </summary>
        /// <param name="reader">The reader to read from.</param>
        /// <param name="context">Context information required for the reading process.</param>
        void Read(IExcelReader reader, IVisualElementContext context);
    }
}
