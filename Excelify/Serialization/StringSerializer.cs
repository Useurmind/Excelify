using BaseLibExt.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelify.Serialization
{
    public class StringSerializer : ISerializer
    {
        public object Deserialize(string value)
        {
            return value.TrimWhitespaces();
        }

        public string Serialize(object value)
        {
            return value.ToString().TrimWhitespaces();
        }
    }
}
