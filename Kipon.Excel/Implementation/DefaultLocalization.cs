using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api.Globalization;

namespace Kipon.Excel.Implementation
{
    internal class DefaultLocalization : Api.Globalization.ILocalization
    {
        string ILocalization.True => "Yes";

        string ILocalization.False => "No";
    }
}
