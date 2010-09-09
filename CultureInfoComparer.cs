
namespace Resx2Xls
{
    using System;
    using System.Collections;
    using System.Globalization;

    /// <summary>
    /// IComparer for the CultureInfo class
    /// </summary>
    public class CultureInfoComparer : IComparer
    {
        /// <summary>
        /// Compares two objects and returns a value indicating whether one is less than, equal to, or greater than the other.
        /// </summary>
        /// <param name="x">The first object to compare.</param>
        /// <param name="y">The second object to compare.</param>
        /// <returns>an int describing which one is greater</returns>
        public int Compare(object x, object y)
        {
            if (((x == null) && (y == null)) || x.Equals(y))
            {
                return 0;
            }

            if (x.Equals(CultureInfo.InvariantCulture) || (y == null))
            {
                return -1;
            }

            if (y.Equals(CultureInfo.InvariantCulture) || (x == null))
            {
                return 1;
            }

            if (!(x is CultureInfo))
            {
                throw new ArgumentException("Can only compare CultureInfo objects.", "x");
            }

            string cxname = ((CultureInfo)x).DisplayName;
            if (!(y is CultureInfo))
            {
                throw new ArgumentException("Can only compare CultureInfo objects.", "y");
            }

            string cyname = ((CultureInfo)y).DisplayName;
            return cxname.CompareTo(cyname);
        }
    }
}
