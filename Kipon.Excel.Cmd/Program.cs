using System;

namespace Kipon.Excel.Cmd
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Here we are");
            var empls = Model.SkatField.FromExcel();

            foreach (var empl in empls)
            {
                Console.WriteLine(empl.Name + " " + empl.Level1 + " " + empl.Level2 + " " + empl.Level3 + " " + empl.Level4 + " " + empl.Description + " " + empl.Subscribe);
            }
        }
    }
}
