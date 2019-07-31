using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api.Globalization
{
    public interface ILocalization
    {
        string True { get; }
        string False { get; }

        string LengthExceededErrorTitle { get; }
        string LengthExceededError(int maxlength);
        string LengthExceededPromptTitle { get; }
        string LengthExceededPrompt(int maxlength);
    }
}
