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
            //The executing directory used to resolve dependent assemblies in the current appDomain
            var executingDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            //The directory where this assembly is located
            var asseblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            if (!executingDir.Equals(asseblyDir, StringComparison.InvariantCultureIgnoreCase))
            {
                AppDomain.CurrentDomain.AssemblyResolve += (object obj, ResolveEventArgs args) =>
                {
                    var currentAssembly = Assembly.GetExecutingAssembly();
                    var assemblyReferences = currentAssembly.GetReferencedAssemblies();

                    var needeAssebly = assemblyReferences.Where(x => x.Name == args.Name).FirstOrDefault();
                    if (needeAssebly == null)
                        return null;

                    //Check the assembly directory for dependencies
                    return Assembly.LoadFrom(asseblyDir + @"\" + args.Name +".dll");
                };
            }
        }

        public void ShowWindow()
        {
            try
            {
                var application = new Application();
                var mainWindow = new MainWindow();
                application.MainWindow = mainWindow;
                application.MainWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                //Handle uncaught exceptions that may have happened in the WPF application. 
            }
        }
    }
}
