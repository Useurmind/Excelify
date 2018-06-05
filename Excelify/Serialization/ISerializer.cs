using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Serialization
{
    public interface ISerializer
    {
        object Deserialize(string value);

        string Serialize(object value);
    }
}
