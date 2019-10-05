using System;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Here we are");

            /*
            var empls = Model.SkatField.FromExcel();

            foreach (var empl in empls)
            {
                Console.WriteLine(empl.Name + " " + empl.Level1 + " " + empl.Level2 + " " + empl.Level3 + " " + empl.Level4 + " " + empl.Description + " " + empl.Subscribe);
            }
            */

            using (var file = new System.IO.FileStream(@"C:\Temp\arbo-phone.xlsx", System.IO.FileMode.Open))
            {
                var phonedata = file.ToObject<Arbodania.Crm.Portal.Model.AppPhoneExport>();
            }
        }
    }
}
