using BaseLibExt.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Serialization
{
    public class IntSerializer : ISerializer
    {
        public object Deserialize(string value)
        {
            return int.Parse(value.TrimWhitespaces());
        }

        public string Serialize(object value)
        {
            return value.ToString();
        }
    }
}
