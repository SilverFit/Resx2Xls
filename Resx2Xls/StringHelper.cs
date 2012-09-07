namespace Resx2Xls
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.IO;
    using System.Linq;
    using System.Text;

    public static class StringHelper
    {
        public static StringCollection ToCollection(string multiline)
        {
            var collection = new StringCollection();
            var lines = GetLines(multiline);
            collection.AddRange(lines.ToArray());
            return collection;
        }

        public static IEnumerable<string> GetLines(string multiLineString)
        {
            var stringReader = new StringReader(multiLineString);
            string line = null;
            while ((line = stringReader.ReadLine()) != null)
            {
                yield return line;
            }
        }

        public static List<string> ListFromCollection(StringCollection collection)
        {
            return GetLines(FromCollection(collection)).ToList();
        }

        public static string FromCollection(StringCollection collection)
        {
            return string.Join(Environment.NewLine, collection.Cast<string>().ToArray());
        }
    }
}
