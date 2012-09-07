namespace Resx2Xls
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;

    /// <summary>
    /// Representation of a .resx file that supports reading comments
    /// </summary>
    public class Resx
    {
        private readonly string path;

        public Resx(string path)
        {
            this.path = path;
        }

        /// <summary>
        /// Initializes a new instance of the Resx class.
        /// </summary>
        /// <param name="path">path of resx file</param>
        public IEnumerable<ResxItem> Read()
        {
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
                    yield return resxitem;
                }
            }
        }

        /// <summary>
        /// Returns true if the .resx file at path is a culture specific resource file
        /// </summary>
        /// <param name="path">path of a .resx file</param>
        /// <returns>true if resource file is culture specific</returns>
        public bool IsCultureSpecific
        {
            get
            {
                FileInfo fi = new FileInfo(path);
                string fname = StripFileExtension(fi.Name);

                string cult = null;
                if (fname.IndexOf(".") != -1)
                {
                    cult = fname.Substring(fname.LastIndexOf('.') + 1);
                }

                if (string.IsNullOrEmpty(cult))
                {
                    return false;
                }

                try
                {
                    var ci = CultureInfo.GetCultureInfo(cult);
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
        
        public string GetRelativePath(string projectRoot)
        {
            return new FileInfo(this.path).FullName.Remove(0, AddBS(projectRoot).Length);
        }

        public string PathWithoutExtension
        {
            get { return StripFileExtension(this.path); }
        }

        /// <summary>
        /// Returns a filepath without the extension
        /// </summary>
        /// <param name="filepath">path of a file</param>
        /// <returns>path of a file without extension</returns>
        private static string StripFileExtension(string filepath)
        {
            var fileinfo = new FileInfo(filepath.Trim());
            return fileinfo.FullName.Substring(0, fileinfo.FullName.Length - fileinfo.Extension.Length);
        }

        /// <summary>
        /// Add a backslash to a path if not present
        /// </summary>
        /// <param name="path">path to add backslash to</param>
        /// <returns>path with backslash</returns>
        private static string AddBS(string path)
        {
            if (path.Trim().EndsWith("\\"))
            {
                return path.Trim();
            }
            else
            {
                return path.Trim() + "\\";
            }
        }
    }
}
