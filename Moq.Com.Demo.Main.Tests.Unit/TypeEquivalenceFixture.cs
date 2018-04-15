using Microsoft.Office.Interop.Excel;
using NUnit.Framework;

namespace Moq.Com.Demo
{
    [TestFixture]
    public class TypeEquivalenceFixture
    {
        [Test]
        public void TypesAreEquivalent()
        {
            var testDllType = typeof(Application);
            var mainDllType = BindingWrapper.WrappedType();

            //  ASSERT --------------------------------------------------------
            Assert.That(mainDllType, Is.EqualTo(testDllType));
        }
    }
}
