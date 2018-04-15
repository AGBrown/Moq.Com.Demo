using System;
using Microsoft.Office.Interop.Excel;

namespace Moq.Com.Demo
{
    public class BindingWrapper
    {
        public static Type WrappedType()
        {
            return typeof(Application);
        }
    }
}
