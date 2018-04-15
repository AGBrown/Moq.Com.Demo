using Microsoft.VisualStudio.OLE.Interop;
using NUnit.Framework;

namespace Moq.Com.Demo
{
    [TestFixture]
    public class TypeEquivalenceFixture
    {
        [Test]
        public void TypesAreEquivalent()
        {
            var testDllType = typeof(IBinding);
            var mainDllType = BindingWrapper.WrappedType();

            //  ASSERT --------------------------------------------------------
            Assert.That(mainDllType, Is.EqualTo(testDllType));
        }
    }
}
