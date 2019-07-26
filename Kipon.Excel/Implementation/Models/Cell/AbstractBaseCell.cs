﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Models
{
    internal abstract class AbstractBaseCell : ICell
    {
        private Coordinate _coordinate;
        private int column;
        private int row;

        internal AbstractBaseCell(int zerobaseColumnIndex, int zerobaseRowIndex)
        {
            this.column = zerobaseColumnIndex;
            this.row = zerobaseRowIndex;
            this._coordinate = new Coordinate(zerobaseColumnIndex, zerobaseRowIndex);
        }

        public Kipon.Excel.Api.ICoordinate Coordinate => this._coordinate;

        public abstract object Value { get; }

        public override string ToString()
        {
            return this.column + ":" + this.row;
        }

        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        public override bool Equals(object obj)
        {
            var other = obj as AbstractBaseCell;

            if (other != null)
            {
                return this.column == other.column && this.row == other.row;
            }
            return false;
        }
    }
}
