﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    internal interface IPopulator
    {
        void Populate(object instance);
    }
}
