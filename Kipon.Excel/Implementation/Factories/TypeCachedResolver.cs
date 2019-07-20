using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    /// <summary>
    /// The purpose of this is to separate resolving actual type of I from populating it.
    /// This will enable caching of needed types, so "expesive analyse of data" based on reflection
    /// only need to be performed once.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <typeparam name="I"></typeparam>
    internal abstract class TypeCachedResolver<T,I,J> : BaseResolver<T,I> 
        where T: class 
        where J:T, IPopulator<I>, new()
    {
        private static Dictionary<string, Type> typeCache = new Dictionary<string, Type>();

        public sealed override T Resolve(I instance)
        {
            var result = base.Resolve(instance);
            if (result != null)
            {
                return result;
            }

            var key = typeof(I).FullName + "|" + typeof(T).FullName;
            if (typeCache.ContainsKey(key))
            {
                var impl = new J();
                impl.Populate(instance);
                return impl;
            }
            else
            {
                var impl = this.ResolveType(instance);
                impl.Populate(instance);
                return impl;
            }
        }


        /// <summary>
        /// Resolve will be called only if we have not seen the instance type I mapped to the return type T before.
        /// </summary>
        /// <typeparam name="J">Return the first time new() of an implementation of T, represented by J</typeparam>
        /// <param name="instance">The instance to resolve from</param>
        /// <returns></returns>
        protected abstract J ResolveType(I instance);
    }
}
