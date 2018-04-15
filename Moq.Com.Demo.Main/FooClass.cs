using Microsoft.Office.Interop.Excel;

namespace Moq.Com.Demo
{
    public interface IFooClass
    {
        string DoWork();
    }

    public class FooClass : IFooClass
    {
        private readonly Application _comBarClass;

        public FooClass(Application comBarClass)
        {
            _comBarClass = comBarClass;
        }

        public string DoWork()
        {
            return _comBarClass.ActivePrinter;
        }
    }
}
