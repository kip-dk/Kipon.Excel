using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Linq
{
    /// <summary>
    /// Single: Public API for transforming an excel stream into an object
    /// Reason: Adding new convinience methods that can read from excel data
    /// </summary>
    public static class LinqReaderInterface
    {
        internal const string USE_ToEnumerable_SUGGESTION = "use ToArray(..) or ToEnumerable(..) or ToList(..) instead";


        #region read from excel
        /// <summary>
        /// Transform an excel stream into a IEnumerable of T instances
        /// 
        /// Based on reflection, this method will transform the data from excel into an enumerable of objects, using
        /// the duck type pattern to resolve maps.
        /// </summary>
        /// <typeparam name="T">The type to resolve</typeparam>
        /// <param name="excel">The excel stream data</param>
        /// <returns>IEnumerable of T </returns>
        public static IEnumerable<T> ToEnumerable<T>(this System.IO.Stream excel, bool mergeAll = false) where T : new()
        {
            var result = new List<T>();
            result.From(excel, mergeAll);
            return result;
        }

        /// <summary>
        /// Transform excel stream into an array of IEnumerable.
        /// This method is simply calling ToArray() on ToEnumerable&lt;T&gt;
        /// </summary>
        /// <typeparam name="T">The type to resolve</typeparam>
        /// <param name="excel">The excel stream data</param>
        /// <returns>array of T</returns>
        public static T[] ToArray<T>(this System.IO.Stream excel, bool mergeAll = false) where T : new()
        {
            return LinqReaderInterface.ToEnumerable<T>(excel, mergeAll).ToArray();
        }

        /// <summary>
        /// Transform excel stream into an array of IEnumerable.
        /// This method is simply calling ToList() on ToEnumerable&lt;T&gt;
        /// </summary>
        /// <typeparam name="T">The type to resolve</typeparam>
        /// <param name="excel">The excel stream data</param>
        /// <returns>List of T</returns>
        public static List<T> ToList<T>(this System.IO.Stream excel, bool mergeAll = false) where T : new()
        {
            return LinqReaderInterface.ToEnumerable<T>(excel, mergeAll).ToList();
        }


        /// <summary>
        /// This method transform an excel into a single object of type T.
        /// This method is typically convinient for transforming workbooks with
        /// more than one sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excel"></param>
        /// <returns></returns>
        public static T ToObject<T>(this System.IO.Stream excel) where T : class
        {
            if (typeof(T).IsArray)
            {
                throw new Kipon.Excel.Exceptions.UnsupportedTypeException(nameof(LinqReaderInterface.ToObject), typeof(T), USE_ToEnumerable_SUGGESTION);
            }

            var parser = new Kipon.Excel.ReaderImplementation.Converters.SpreadsheetConverter();
            return parser.Convert<T>(excel);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <param name="excel"></param>
        /// <returns></returns>
        public static void From<T>(this T t, System.IO.Stream excel, bool mergeAll = false)
        {
            var parser = new Kipon.Excel.ReaderImplementation.Converters.SpreadsheetConverter();
            parser.ConvertInto(t, excel, mergeAll);
        }
        #endregion

    }
}
