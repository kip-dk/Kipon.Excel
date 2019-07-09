using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Serialization
{
    public class DefaultExcelSerializer<T> : IExcelSerializable<T> where T: new()
    {
        public T Deserialize(Stream input)
        {
            throw new NotImplementedException();
        }

        public void Serialize(Stream output, T data)
        {
            throw new NotImplementedException();
        }
    }
}
