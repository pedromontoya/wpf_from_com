using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using WPF.Sample;

namespace COM
{
    [Guid("EE942FAD-D852-4E53-9113-33CFB339769C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("WPF.FROM.COM")]
    public class SampleWpfInvoker: IWpfInvoker
    {
        public SampleWpfInvoker()
        {
            var executingDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            var asseblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            if (!executingDir.Equals(asseblyDir, StringComparison.InvariantCultureIgnoreCase))
            {
                AppDomain.CurrentDomain.AssemblyResolve += delegate(object obj, ResolveEventArgs e)
                {

                };
            }
        }
        //mainDir = appdomain.currentDom.settupInfo.appbase
        //dotNetPath = path.getDirName(Assembly.GetExecutingAssem().Location)

        //if(!mainDir.equal(dotNetPath)
        //Appdomain.currentasembly.resolve +=
        public void ShowWindow()
        {
            var application = new Application();
            var mainWindow = new MainWindow();
            application.MainWindow = mainWindow;

            application.MainWindow.ShowDialog();
        }
    }
}
