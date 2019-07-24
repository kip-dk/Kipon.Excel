﻿using Kipon.Excel.Api;
using Kipon.Excel.Implementation.Models.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    internal class CellsResolver : TypeCachedResolver<IEnumerable<ICell>, Models.Cells.AbstractBaseCells>
    {
        private static readonly Type[] SUPPORTED_SHEET_PROPERTY_TYPES = new Type[]
        {
            typeof(System.Int16),
            typeof(System.Int32),
            typeof(System.Int64),
            typeof(System.IntPtr),
            typeof(System.UInt16),
            typeof(System.UInt32),
            typeof(System.UInt64),
            typeof(System.UIntPtr),
            typeof(System.Boolean),
            typeof(System.Decimal),
            typeof(System.Double),
            typeof(System.Enum),
            typeof(System.String),
            typeof(System.DateTime)
        };

        protected override AbstractBaseCells ResolveType(Type instanceType)
        {
            throw new NotImplementedException();
        }

        internal bool HasCells(Type type)
        {
            return true;
        }
    }
}