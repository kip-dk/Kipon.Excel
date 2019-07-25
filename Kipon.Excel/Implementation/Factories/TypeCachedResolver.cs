using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    /// <summary>
    /// The purpose of this is to separate resolving actual type T from I from populating the result T with values from I.
    /// This will enable caching of needed types, so "expesive analyse of data" ex. based on reflection
    /// only need to be performed once.
    /// </summary>
    /// <typeparam name="T">the interface to resolve to</typeparam>
    /// <typeparam name="I">The instance to resolve from</typeparam>
    /// <typeparam name="J">J is the actual implementation of T. It need to impl.IPopolator to allow the typed based cache of T to parse the data provided, and it need to have a public constructor to allow the cache to create an instance</typeparam>
    internal abstract class TypeCachedResolver<T,J> : AbstractBaseResolver<T> 
        where T: class 
        where J:T, IPopulator
    {
        private static Dictionary<string, System.Reflection.ConstructorInfo> typeCache = new Dictionary<string, System.Reflection.ConstructorInfo>();

        public sealed override T Resolve(object instance)
        {
            var result = base.Resolve(instance);
            if (result != null)
            {
                return result;
            }

            var key = instance.GetType().FullName + "|" + typeof(T).FullName;
            if (typeCache.ContainsKey(key))
            {
                var impl = typeCache[key].Invoke(new object[0]) as IPopulator;

                impl.Populate(instance);
                return (T)impl;
            }
            else
            {
                var impl = this.ResolveType(instance.GetType());
                typeCache.Add(key, impl.GetType().GetConstructor(new Type[0]));

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
        protected abstract J ResolveType(Type instanceType);
    }
}
