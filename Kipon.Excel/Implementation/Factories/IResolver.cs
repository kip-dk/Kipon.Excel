using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    /// <summary>
    // Interface to resolve an object instance to its needed implementation T. 
    /// Classes impl. this interface must resolve a type T from the instance I, and populate
    /// all information in the result prior to return based on data int the input I.
    /// </summary>
    /// <typeparam name="T">The expected return type</typeparam>
    internal interface IResolver<T>  where T : class
    {
        T Resolve<I>(I instance);
    }
}
