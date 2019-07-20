using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    /// <summary>
    // Interface to resolve an instance I to its needed implementation T. 
    /// Classes impl. this interface must resolve a type T from the instance I, and populate
    /// all information in the result prior to return.
    /// </summary>
    /// <typeparam name="T">The expected return type</typeparam>
    /// <typeparam name="I">The instance that need to be transformed into T somehow</typeparam>
    internal interface IResolver<T, I>  where T : class
    {
        T Resolve(I instance);
    }
}
