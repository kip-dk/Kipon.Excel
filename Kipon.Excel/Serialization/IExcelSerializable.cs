using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Serialization
{
    public interface IExcelSerializable<T> where T: new()
    {
        void Serialize(System.IO.Stream output, T data);
        T Deserialize(System.IO.Stream input);
    }
}
