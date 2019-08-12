using System;

namespace Kipon.Excel.Cmd
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Here we are");
            var empls = Model.Employee.FromExcel();

            foreach (var empl in empls)
            {
                Console.WriteLine(empl.Name + " " + empl.Mobile);
            }
        }
    }
}
