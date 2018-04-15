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
            var testDllType2 = typeof(Application);
            var mainDllType = BindingWrapper.WrappedType();

            //  ASSERT --------------------------------------------------------
            Assert.Multiple(() => {
                Assert.That(testDllType, Is.SameAs(testDllType2));
                Assert.That(testDllType, Is.SameAs(mainDllType));
                Assert.That(testDllType, Is.EqualTo(mainDllType));
            });
        }
    }
}
