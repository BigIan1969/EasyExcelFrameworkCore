using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcelFramework
{
    internal class StringConverter
    {
        public object DetectType(string stringValue)
        {
            if (stringValue == null)
                return null;
            var expectedTypes = new List<Type> { typeof(DateTime)};
            foreach (var type in expectedTypes)
            {
                TypeConverter converter = TypeDescriptor.GetConverter(type);
                if (converter.CanConvertFrom(typeof(string)))
                {
                    try
                    {
                        // You'll have to think about localization here
                        object newValue = converter.ConvertFromInvariantString(stringValue);
                        if (newValue != null)
                        {
                            return newValue;
                        }
                    }
                    catch
                    {
                        // Can't convert given string to this type
                        continue;
                    }

                }

            }
            if (stringValue[0].Equals("\"") & stringValue[stringValue.Length-1].Equals("\""))
            {
                return stringValue;
            }
            if (int.TryParse(stringValue, out int result))
            {
                return result;
            }
            if (float.TryParse(stringValue, out float floatresult))
            {
                return floatresult;
            }
            return stringValue;
        }
    }
}
