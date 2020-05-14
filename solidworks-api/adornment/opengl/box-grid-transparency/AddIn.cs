using CodeStack.SwEx.AddIn;
using CodeStack.SwEx.AddIn.Attributes;
using System;
using System.Runtime.InteropServices;

namespace CodeStack.OpenGlBoxGrid
{
    [ComVisible(true), Guid("FAB0F03B-785E-4113-B120-E18E7C73B9EB")]
    [AutoRegister("OpenGL Box Grid")]
    public class AddIn : SwAddInEx
    {
        public override bool OnConnect()
        {
            CreateDocumentsHandler<OpenGlDocumentHandler>();
            return true;
        }
    }
}
