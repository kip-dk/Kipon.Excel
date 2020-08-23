using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Attributes
{
    /// <summary>
    /// Declare a property to be optional, so a match between a model and a defacto sheet is true, even if the column is not present in the sheet
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class OptionalAttribute : Attribute
    {
    }
}
