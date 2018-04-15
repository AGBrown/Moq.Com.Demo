using System;
using Autofac.Extras.Moq;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;

namespace Moq.Com.Demo
{
    [TestFixture] public class FooClassFixture
    {
        public enum UseMainTestInput { UseMain, UseTest }
        public enum AssertionTestInput { AssertActual, VerifyMock }

        [Test]
        public void DoWork(
            [Values] UseMainTestInput useTypeFrom, 
            [Values] AssertionTestInput assertion)
        {
            if (useTypeFrom == UseMainTestInput.UseMain)
                DoWork<FooClass>(assertion);
            else
                DoWork<FooClass<Application>>(assertion);
        }

        public void DoWork<T>(AssertionTestInput assertion)
            where T : IFooClass
        {
            //    NUnit test, edited for brevity
            using (var mockFactory = AutoMock.GetLoose())
            {
                var expected = Guid.NewGuid().ToString();
                mockFactory.Mock<Application>()
                           .SetupGet(x => x.ActivePrinter)
                           .Returns(expected);

                var sut = mockFactory.Create<T>();

                var actual = sut.DoWork();

                //  AssertActual & ExpectFalse will always pass as mocks return false even without setups
                if (assertion == AssertionTestInput.AssertActual)
                    Assert.That(actual, Is.EqualTo(expected));
                else
                    mockFactory.Mock<Application>().VerifyGet(x => x.ActivePrinter, Times.Once);
            }
        }
    }
}
