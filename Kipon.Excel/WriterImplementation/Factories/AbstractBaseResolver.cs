using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Factories
{
    /// <summary>
    /// Base class for resolving types, from an instance a type T must be resolved
    /// If instance is already a T, it will simply return instance as T, otherwise default(T) to indicate
    /// the resolving instance was not possible
    /// </summary>
    /// <typeparam name="T">An implemenation of T</typeparam>

    internal abstract class AbstractBaseResolver<T> : IResolver<T> where T : class
    {
        public virtual T Resolve<I>(I instance) 
        {
            if (instance == null)
            {
                throw new Kipon.Excel.Exceptions.NullInstanceException(typeof(T));
            }

            if (instance is T)
            {
                return instance as T;
            }

            return default(T);
        }
    }
}
