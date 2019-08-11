using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Cmd
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0 || (args.Length == 1 && args[0] == "export"))
            {
                Kipon.Excel.Cmd.Export.ExportTest.Run();
                return;
            }

            if (args.Length == 1 && args[0] == "import")
            {
                Kipon.Excel.Cmd.Import.ImportTest.Run();
            }
        }
    }
}
