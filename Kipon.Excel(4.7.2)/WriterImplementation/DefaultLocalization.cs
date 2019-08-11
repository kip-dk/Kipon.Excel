using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api.Globalization;

namespace Kipon.Excel.WriterImplementation
{
    internal class DefaultLocalization : Api.Globalization.ILocalization
    {
        string ILocalization.True => "Yes";

        string ILocalization.False => "No";

        string ILocalization.LengthExceededErrorTitle => "Length exceeded";

        string ILocalization.LengthExceededPromptTitle => "Max length";

        string ILocalization.LengthExceededError(int maxlength)
        {
            return $"This should contain a string of max length {maxlength}";
        }

        string ILocalization.LengthExceededPrompt(int maxlength)
        {
            return $"Max length {maxlength}";
        }
    }
}
