﻿using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    /// <summary>
    /// From an instance I, create an implementaton J that implements T, where T is IEnumerable<ISheet>
    /// </summary>
    /// <typeparam name="T">T is the interface to be resolved</typeparam>
    /// <typeparam name="I">I is the instance to resolve from</typeparam>
    /// <typeparam name="J">J is the actual implementation of T. It must have a constructor that take 0 arguments</typeparam>
    internal class SheetsResolver<T,J> : TypeCachedResolver<IEnumerable<ISheet>, Models.Sheets>
        where J : T, IPopulator
    {
        protected override Models.Sheets ResolveType(object instance) 
        {
            var type = instance.GetType();
            if (type.IsArray)
            {
                var elementType = type.GetElementType();
                if (elementType is Kipon.Excel.Api.ISheet)
                {
                    return new Models.Sheets();
                }
            }

            return null;
        }
    }
}