using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.WriterImplementation.Factories
{
    public class AbstractBaseResolverTest
    {
        [Test]
        public void ResolveTest()
        {
            // Test than a List impl. of the ienumable request will return the instance it self
            {
                var impl = new List<Kipon.Excel.Api.ISheet>();
                var resolver = new Resolver<IEnumerable<Kipon.Excel.Api.ISheet>>();
                var other = resolver.Resolve(impl);
                Assert.AreEqual(impl, other);
            }

            // Test that an array impl. of the ienuable request iill rturn the instance it self
            {
                var impl = new Kipon.Excel.Api.ISheet[0];
                var resolver = new Resolver<IEnumerable<Kipon.Excel.Api.ISheet>>();
                var other = resolver.Resolve(impl);
                Assert.AreEqual(impl, other);
            }

            // Test that resolve something that does not math the resolver interface will rturn null
            {
                var impl = new List<string>();

                var resolver = new Resolver<IEnumerable<Kipon.Excel.Api.ISheet>>();
                var other = resolver.Resolve(impl);
                Assert.IsNull(other);
            }
        }


        internal class Resolver<T> : Kipon.Excel.WriterImplementation.Factories.AbstractBaseResolver<T> where T: class
        {
        }
    }
}
