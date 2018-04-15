using Microsoft.Office.Interop.Excel;

namespace Moq.Com.Demo
{
    public class FooClass<T> : FooClass where T : Application
    {
        public FooClass(T comBarClass) : base(comBarClass) { }
    }
}
