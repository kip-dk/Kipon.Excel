using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Attributes;

namespace Arbodania.Crm.Portal.Model
{
    public class AppPhoneExport
    {
        [Sheet(1)]
        public Phone[] Phones { get; set; }

        [Sheet(2)]
        public User[] Users { get; set; }

        [Sheet(3)]
        public PhoneUser[] PhoneUsers { get; set; }

        [Sheet(4)]
        public PhoneSupplier[] PhoneSuppliers { get; set; }

        [Sheet(5)]
        public Supplier[] Suppliers { get; set; }

        public class Phone
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public int Environment { get; set; }
            public string Email { get; set; }
        }

        public class User
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public string Email { get; set; }
            public string Permissions { get; set; }

            [Kipon.Excel.Attributes.Ignore]
            public string _permissions
            {
                get; set;
            }
        }

        public class PhoneUser
        {
            public Guid Id { get; set; }
            public Guid UserId { get; set; }
            public Guid PhoneId { get; set; }
        }

        public class PhoneSupplier
        {
            public Guid Id { get; set; }
            public Guid SupplierId { get; set; }
            public Guid PhoneId { get; set; }
        }

        public class Supplier
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
        }
    }
}
