using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api
{
    /// <summary>
    /// Responsibiliy: Represent the position in the excel coordinate system
    /// Reason: ?
    /// </summary>
    public interface ICoordinate
    {
        /// <summary>
        /// returns a list of points to return a unique position in a multi dimensional coordinate system
        /// The method must return same point representation in same position for the full implementation, 
        /// ex. for a 2 dimentional sheet, the x axis as first, and the y axis for second. If the sheet support
        /// more than two dimensions, this will simply be represented with the 3 element in the list z.
        /// </summary>
        IEnumerable<int> Point { get; }
    }
}
