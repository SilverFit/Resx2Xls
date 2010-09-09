
namespace Resx2Xls
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Xml.Linq;

    /// <summary>
    /// Representation of a .resx file that supports reading comments
    /// </summary>
    public class Resx
    {
        /// <summary>
        /// Initializes a new instance of the Resx class.
        /// </summary>
        /// <param name="path">path of resx file</param>
        public static IEnumerable<ResxItem> Read(string path)
        {
            List<ResxItem> rows = new List<ResxItem>();
            XDocument resxXML = XDocument.Load(path);
            foreach (var row in resxXML.Root.Descendants("data"))
            {
                var name = row.Attribute("name");
                var type = row.Attribute("type");
                var value = row.Descendants("value").FirstOrDefault();
                var comment = row.Descendants("comment").FirstOrDefault();
                // Only read if type is null : string
                if (name != null && type == null)
                {
                    var resxitem = new ResxItem(
                        name.Value,
                        value == null ? string.Empty : value.Value,
                        comment == null ? string.Empty : comment.Value);
                    rows.Add(resxitem);
                }
            }
            return rows;
        }
    }
}
