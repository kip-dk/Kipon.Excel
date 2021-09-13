using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Api
{
    public interface ISheetProtectionProperties
    {
        string AlgorithmName { get; }
        string HashValue { get; }
        string SaltValue { get; }
        System.UInt32 SpinCount { get; }
    }
}
