using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Attributes;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd.Model
{
    public class SkatField
    {

        public static SkatField[] FromExcel()
        {
            using (var fs = new System.IO.FileStream(@"C:\Temp\atp.xlsx", System.IO.FileMode.Open))
            {
                return fs.ToArray<SkatField>();
            }
        }


        [Title("N1")]
        public string Level1 { get; set; }

        [Title("N2")]
        public string Level2 { get; set; }

        [Title("N3")]
        public string Level3 { get; set; }

        [Title("N4")]
        public string Level4 { get; set; }

        [Title("Felt")]
        public string FieldId { get; set; }

        [Title("Navn")]
        public string Name { get; set; }

        [Title("Beskrivelse")]
        public string Description { get; set; }

        [Title("Supplerende oplysninger")]
        public string SupplInformation { get; set; }

        [Title("Abonnement")]
        public string Subscribe { get; set; }

        [Title("Typenr i indberetning")]
        public string TypeNumber { get; set; }

        [Title("Datatype i XML")]
        public string Datatype { get; set; }

        [Title("Format")]
        public string Format { get; set; }
    }
}
