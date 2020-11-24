using System;
using System.Activities;
using System.Activities.XamlIntegration;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Run("Test1.xaml");

        }
        public static void Run(string fileName)
        {
            var workflow = ActivityXamlServices.Load(fileName, new ActivityXamlServicesSettings() { CompileExpressions = true });
            WorkflowInvoker.Invoke(workflow);
        }
    }
}
