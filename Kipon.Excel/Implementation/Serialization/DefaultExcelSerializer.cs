using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Serialization
{
    internal class DefaultExcelSerializer<T> : IExcelSerializable<T> where T: new()
    {
        T IExcelSerializable<T>.Deserialize(Stream input)
        {
            throw new NotImplementedException();
        }

        void IExcelSerializable<T>.Serialize(Stream output, T data)
        {
            throw new NotImplementedException();
        }
    }
}
