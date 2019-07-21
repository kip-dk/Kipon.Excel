using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    /// <summary>
    /// Base class for resolving types, where I actualy is an instance of T, it will simply return I as a T
    /// If I is not a T, default(T) will be returned, indicating that it was unable to do anything relevant
    /// </summary>
    /// <typeparam name="T">An implemenation of T</typeparam>
    /// <typeparam name="I">The instance to resolve from</typeparam>

    internal abstract class AbstractBaseResolver<T> : IResolver<T> where T : class
    {
        public virtual T Resolve(object instance) 
        {
            if (instance is T)
            {
                return instance as T;
            }

            return default(T);
        }
    }
}
