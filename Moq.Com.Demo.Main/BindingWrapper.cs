using System;
using Microsoft.VisualStudio.OLE.Interop;

namespace Moq.Com.Demo
{
    public class BindingWrapper
    {
        public static Type WrappedType()
        {
            return typeof(IBinding);
        }
    }
}
