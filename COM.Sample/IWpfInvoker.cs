using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace COM
{
    [Guid("2AB00BE3-5608-47BC-94B9-BFC18F3E8C68")]
    public interface IWpfInvoker
    {
        void ShowWindow();
    }
}
