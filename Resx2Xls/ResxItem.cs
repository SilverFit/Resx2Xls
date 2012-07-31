
namespace Resx2Xls
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Row item in resx file
    /// </summary>
    public class ResxItem
    {
        /// <summary>
        /// Key of a Resx row
        /// </summary>
        public readonly string Key;

        /// <summary>
        /// Value of a Resx row
        /// </summary>
        public readonly string Value;

        /// <summary>
        /// Comment of a Resx row
        /// </summary>
        public readonly string Comment;

        /// <summary>
        /// Initializes a new instance of the ResxItem class.
        /// </summary>
        /// <param name="key">Key of a Resx row</param>
        /// <param name="value">Value of a Resx row</param>
        /// <param name="comment">Comment of a Resx row</param>
        public ResxItem(string key, string value, string comment)
        {
            this.Key = key;
            this.Value = value;
            this.Comment = comment;
        }

        /// <summary>
        /// Returns an empty ResxItem
        /// </summary>
        public static ResxItem Empty
        {
            get
            {
                return new ResxItem(string.Empty, string.Empty, string.Empty);
            }
        }
    }
}
