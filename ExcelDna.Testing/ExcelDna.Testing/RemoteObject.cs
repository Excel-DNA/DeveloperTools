using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;
using System.Runtime.InteropServices;

namespace ExcelDna.Testing
{
    public class RemoteObject : LongLivedMarshalByRefObject
    {
        public SerializableRunSummary RunTests(string testAssemblyPath, string testAssemblyConfigurationFile, ExcelTestCase[] testCases, IMessageSink diagnosticMessageSink, IMessageSink executionMessageSink, ITestFrameworkExecutionOptions executionOptions, IMessageBus messageBus)
        {
            var testAssembly = new TestAssembly(new ReflectionAssemblyInfo(Assembly.LoadFrom(testAssemblyPath)), testAssemblyConfigurationFile);
            RemoteTestAssemblyRunner runner = new RemoteTestAssemblyRunner(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions, messageBus);
            return runner.Run();
        }

        public void CloseHost()
        {
            const uint WM_CLOSE = 0x0010;
            PostMessage(ExcelDna.Integration.ExcelDnaUtil.WindowHandle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    }
}
