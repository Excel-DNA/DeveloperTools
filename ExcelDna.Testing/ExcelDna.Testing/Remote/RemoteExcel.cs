using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Xunit.Sdk;
using System.Runtime.InteropServices;

namespace ExcelDna.Testing.Remote
{
    public interface IRemoteExcel
    {
        event EventHandler<BusMessageEventArgs> BusMessage;

        Task<SerializableRunSummary> RunTestsAsync(string testAssemblyPath, string testAssemblyConfigurationFile, string[] testCases);
        Task CloseHostAsync();
    }

    public class RemoteExcel : IRemoteExcel
    {
        public event EventHandler<BusMessageEventArgs> BusMessage;

        public async Task<SerializableRunSummary> RunTestsAsync(string testAssemblyPath, string testAssemblyConfigurationFile, string[] testCases)
        {
            var testAssembly = new TestAssembly(new ReflectionAssemblyInfo(Assembly.LoadFrom(testAssemblyPath)), testAssemblyConfigurationFile);
            MessageBusProxy messageBusProxy = new MessageBusProxy(SendBusMessage);
            RemoteTestAssemblyRunner runner = new RemoteTestAssemblyRunner(testAssembly, testCases.Select(i => ExcelTestCase.DeserializeFromString(i)), null, null, new TestFrameworkOptions(), messageBusProxy);
            return await runner.RunAsync();
        }

#pragma warning disable CS1998
        public async Task CloseHostAsync()
        {
            const uint WM_CLOSE = 0x0010;
            PostMessage(ExcelDna.Integration.ExcelDnaUtil.WindowHandle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
#pragma warning restore CS1998

        private bool SendBusMessage(Xunit.Abstractions.IMessageSinkMessage message)
        {
            if (BusMessage == null)
                return false;

            BusMessage?.Invoke(null, new BusMessageEventArgs(message));
            return true;
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    }
}
