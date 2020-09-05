using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;
using System.Runtime.Serialization.Formatters;
using System.Threading;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    internal class ExcelTestAssemblyRunner : XunitTestAssemblyRunner
    {
        private readonly ITestAssembly testAssembly;
        private readonly IMessageSink diagnosticMessageSink;
        private readonly IMessageSink executionMessageSink;
        private readonly ITestFrameworkExecutionOptions executionOptions;
        private readonly ExcelRunner excelRunner;

        public ExcelTestAssemblyRunner(ITestAssembly testAssembly, IEnumerable<IXunitTestCase> testCases, IMessageSink diagnosticMessageSink, IMessageSink executionMessageSink, ITestFrameworkExecutionOptions executionOptions)
            : base(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions)
        {
            this.testAssembly = testAssembly;
            this.diagnosticMessageSink = diagnosticMessageSink;
            this.executionMessageSink = executionMessageSink;
            this.executionOptions = executionOptions;
            excelRunner = new ExcelRunner();
        }

        protected override async Task<RunSummary> RunTestCollectionsAsync(IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            IEnumerable<IXunitTestCase> localTestCases = TestCases.Except(TestCases.OfType<ExcelTestCase>());
            IEnumerable<ExcelTestCase> excelTestCases = TestCases.OfType<ExcelTestCase>().Where(i => !i.UseCOM);
            IEnumerable<ExcelTestCase> excelCOMTestCases = TestCases.OfType<ExcelTestCase>().Where(i => i.UseCOM);

            var result = await LocalRunTestCasesAsync(localTestCases, messageBus, cancellationTokenSource);
            if (excelCOMTestCases.Count() > 0)
                result.Aggregate(await COMRunTestCasesAsync(excelCOMTestCases, messageBus, cancellationTokenSource));
            if (excelTestCases.Count() > 0)
                result.Aggregate(await RemoteRunTestCasesAsync(excelTestCases, messageBus, cancellationTokenSource));
            return result;
        }

        private async Task<RunSummary> LocalRunTestCasesAsync(IEnumerable<IXunitTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            var allTestCases = TestCases;
            TestCases = testCases;
            var result = await base.RunTestCollectionsAsync(messageBus, cancellationTokenSource);
            TestCases = allTestCases;
            return result;
        }

        private async Task<RunSummary> RemoteRunTestCasesAsync(IEnumerable<ExcelTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            RunSummary result = new RunSummary();
            try
            {
                ExcelStartupEvent.Create();
                Process excelProcess = excelRunner.Start(testAssembly.Assembly.AssemblyPath);
                if (!ExcelStartupEvent.Wait(30000))
                    throw new System.ApplicationException("Excel startup failed.");

                channel = RegisterIpcChannel("ExcelDna.Testing.ClientChannel", Guid.NewGuid().ToString(), false);
                RemoteObject remoteObject = (RemoteObject)RemotingServices.Connect(typeof(RemoteObject), "ipc://xxx1000/RemoteObject.rem");
                RemoteTestAssemblyRunner remoteTestAssemblyRunner = CreateRemoteTestAssemblyRunner(testCases, remoteObject, cancellationTokenSource, messageBus);

                if (Debugger.IsAttached)
                    VS.VisualStudioInstance.AttachDebugger(excelProcess);

                result = remoteTestAssemblyRunner.Run();
                remoteObject.CloseHost();
            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }

            return result;
        }

        private async Task<RunSummary> COMRunTestCasesAsync(IEnumerable<ExcelTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            try
            {
                COMUtil.SetApplication(new Microsoft.Office.Interop.Excel.Application());
                COMUtil.Application.RegisterXLL(ExcelRunner.GetXllPath(testAssembly.Assembly.AssemblyPath, Marshal.SizeOf(COMUtil.Application.HinstancePtr) == 8 ? Bitness.Bit64 : Bitness.Bit32));
            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
                return new RunSummary();
            }

            TestCases = testCases;
            return await base.RunTestCollectionsAsync(messageBus, cancellationTokenSource);
        }

        IpcChannel channel;

        public static IpcChannel RegisterIpcChannel(string name, string portName, bool ensureSecurity)
        {
            var ipcChannel = new IpcChannel(
                properties: new Hashtable
                {
                    ["name"] = name,
                    ["portName"] = portName,
                },
                clientSinkProvider: new BinaryClientFormatterSinkProvider { },
                serverSinkProvider: new BinaryServerFormatterSinkProvider { TypeFilterLevel = TypeFilterLevel.Full });

            ChannelServices.RegisterChannel(ipcChannel, ensureSecurity);
            return ipcChannel;
        }

        private RemoteTestAssemblyRunner CreateRemoteTestAssemblyRunner(IEnumerable<ExcelTestCase> testCases, RemoteObject remoteObject, CancellationTokenSource cancellationTokenSource, IMessageBus messageBus)
        {
            var diagnosticSink = new DynMessageSink(diagnosticMessageSink, message =>
            {
                if (cancellationTokenSource.Token.IsCancellationRequested)
                    return false;

                return diagnosticMessageSink.OnMessage(message);
            });

            var bus = new DynMessageBus(messageBus, message =>
            {
                if (cancellationTokenSource.Token.IsCancellationRequested)
                    return false;

                switch (message)
                {
                    case ITestAssemblyStarting assemblyStarting:
                    case ITestAssemblyFinished assemblyFinished:
                        return true;
                    case ITestCaseStarting testCaseStarting:
                        break;
                    case ITestCaseFinished testCaseFinished:
                        break;
                }

                return messageBus.QueueMessage(message);
            });

            return remoteObject.CreateTestAssemblyRunner(testAssembly.Assembly.AssemblyPath, testAssembly.ConfigFileName, testCases.ToArray(), diagnosticSink, null, executionOptions, bus);
        }

        private class DynMessageSink : LongLivedMarshalByRefObject, IMessageSink
        {
            private readonly IMessageSink messageSink;

            public DynMessageSink(IMessageSink messageSink, Func<IMessageSinkMessage, bool> onMessage)
            {
                this.messageSink = messageSink;
                OnMessageCallback = onMessage;
            }

            public Func<IMessageSinkMessage, bool> OnMessageCallback { get; }

            public bool OnMessage(IMessageSinkMessage message)
                => OnMessageCallback(message);
        }

        private class DynMessageBus : LongLivedMarshalByRefObject, IMessageBus
        {
            private readonly IMessageBus messageBus;

            public DynMessageBus(IMessageBus messageSink, Func<IMessageSinkMessage, bool> onMessage)
            {
                this.messageBus = messageSink;
                OnMessageCallback = onMessage;
            }

            public Func<IMessageSinkMessage, bool> OnMessageCallback { get; }

            public void Dispose()
            {
            }

            public bool QueueMessage(IMessageSinkMessage message)
                => OnMessageCallback(message);
        }

    }
}