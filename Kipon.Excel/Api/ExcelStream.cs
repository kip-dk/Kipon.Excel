using System;
using System.Collections.Generic;
using System.Text;
using Kipon.Excel.Linq;
using System.Linq;

namespace Kipon.Excel.Api
{
    /// <summary>
    /// A stream context. This part of the api is very experimental.
    /// </summary>
    public class ExcelStream : IDisposable 
    {
        private string filename;
        private System.IO.Stream stream;

        internal ExcelStream()
        {
        }

        public ExcelStream(string filename)
        { 
            if (string.IsNullOrEmpty(filename))
            {
                throw new ArgumentException("filename cannot be null or empty");
            }
            this.filename = filename;
        }

        public ExcelStream(System.IO.Stream stream)
        { 
            if (stream == null)
            {
                throw new ArgumentException("stream cannot be null");
            }
            this.stream = stream;
        }

        /// <summary>
        /// Set this property with your own logger before calling the From/To excel method to get logging details
        /// </summary>
        public ILog Log { get; set; }

        public bool MergeAll { get; set; } = false;

        public IEnumerable<T> ToEnumerable<T>() where T : new()
        {
            if (!string.IsNullOrEmpty(filename))
            {
                this.stream = new System.IO.FileStream(filename, System.IO.FileMode.Open);
            }

            var result = new List<T>();
            var parser = new Kipon.Excel.ReaderImplementation.Converters.SpreadsheetConverter(this);
            parser.ConvertInto(result, this.stream, this.MergeAll);
            return result.ToArray();
        }

        public T[] ToArray<T>() where T : new ()
        {
            return this.ToEnumerable<T>().ToArray();
        }

        public List<T> ToList<T>() where T : new()
        {
            return this.ToEnumerable<T>().ToList();
        }

        public void Dispose()
        {
            if (!string.IsNullOrEmpty(this.filename) && this.stream != null)
            {
                this.stream.Close();
            }
        }
    }
}
