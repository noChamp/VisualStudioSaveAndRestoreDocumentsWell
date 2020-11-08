using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentWellLoader
{
    public static class SimpleContainer
    {
        public static IVsUIShellDocumentWindowMgr WinManager { get; set; }

        public static IStream Stream { get; set; }
    }
}
