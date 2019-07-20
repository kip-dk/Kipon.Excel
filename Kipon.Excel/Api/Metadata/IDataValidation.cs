using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api.Metadata
{
    internal interface IDataValidation
    {
        DataValidationValues Type { get; }
        DataValidationOperatorValues Operator { get; }
        bool AllowBlank { get; }
        bool ShowInputMessage { get; }
        bool ShowErrorMessage { get; }
        string ErrorTitle { get; }
        string Error { get; }
        string PromptTitle { get; }
        string Prompt { get; }
#warning, validate if on rule can contain more than on range
        Implementation.OpenXml.Types.Range SequenceOfReferences { get; }

    }
}
