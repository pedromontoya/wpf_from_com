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

            //Assembly dependencies are resolved based on the appDomain application base, this will very based on 
            //where the COM accessible assembly is invoked from. The fix for this is to register a delegate with 
            //the assemblyResolve event on the appDomain. This allows you to search for dependent assemblies in other
            //directories. 
            if (!executingDir.Equals(asseblyDir, StringComparison.InvariantCultureIgnoreCase))
            {
                AppDomain.CurrentDomain.AssemblyResolve += (object obj, ResolveEventArgs args) =>
                {
                    //This simple implementation just looks for dependent assemblies in the list of those
                    //Referenced by the current executing assembly. 
                    var currentAssembly = Assembly.GetExecutingAssembly();
                    var assemblyReferences = currentAssembly.GetReferencedAssemblies();

                    var needeAssebly = assemblyReferences.Where(x => x.Name == args.Name).FirstOrDefault();
                    if (needeAssebly == null)
                        return null;

                    //If the referenced assembly is in the 'assemblyDir', return it
                    var assemblyPath = asseblyDir + @"\" + args.Name +".dll";
                    if (File.Exists(assemblyPath))
                    {
                        return Assembly.LoadFrom(assemblyPath);
                    }

                    //Still not found
                    return null;
                };
            }
        }

        public void ShowWindow()
        {
            try
            {
                //Create the windows application for this AppDomain
                var application = new Application();
                
                //Specify the main window of the Application
                var mainWindow = new MainWindow();
                application.MainWindow = mainWindow;

                //Show the main window as a dialog. This will keep the application running until
                //the open window is closed.
                application.MainWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                //Handle uncaught exceptions that may have happened in the WPF application. 
            }
        }
    }
}
