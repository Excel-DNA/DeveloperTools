using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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

        public ExcelTestAssemblyRunner(ITestAssembly testAssembly, IEnumerable<IXunitTestCase> testCases, IMessageSink diagnosticMessageSink, IMessageSink executionMessageSink, ITestFrameworkExecutionOptions executionOptions)
            : base(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions)
        {
            this.testAssembly = testAssembly;

            this.diagnosticMessageSink = diagnosticMessageSink;
            this.executionMessageSink = executionMessageSink;
            this.executionOptions = executionOptions;
        }

        protected override async Task<RunSummary> RunTestCollectionsAsync(IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            IEnumerable<IXunitTestCase> localTestCases = TestCases.Except(TestCases.OfType<ExcelTestCase>());
            IEnumerable<ExcelTestCase> excelTestCases = TestCases.OfType<ExcelTestCase>();

            var result = await Local_RunTestCasesAsync(localTestCases, messageBus, cancellationTokenSource);
            if (excelTestCases.Count() > 0)
                result.Aggregate(await Remote_RunTestCasesAsync(excelTestCases, messageBus, cancellationTokenSource));
            return result;
        }

        private async Task<RunSummary> Local_RunTestCasesAsync(IEnumerable<IXunitTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            var allTestCases = TestCases;
            TestCases = testCases;
            var result = await base.RunTestCollectionsAsync(messageBus, cancellationTokenSource);
            TestCases = allTestCases;
            return result;
        }

        private async Task<RunSummary> Remote_RunTestCasesAsync(IEnumerable<ExcelTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            RunSummary result = new RunSummary();
            try
            {
                Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", "\"" + @"E:\programming\Govert\Excel-DNA Developer Tools\ExcelDna.Testing\Examples\bin\Debug\Examples-AddIn64.xll" + "\"");
                Thread.Sleep(6000);

                channel = RegisterIpcChannel("ExcelDna.Testing.ClientChannel", Guid.NewGuid().ToString(), false);


                RemoteObject remoteObject = (RemoteObject)RemotingServices.Connect(typeof(RemoteObject), "ipc://xxx1000/RemoteObject.rem");

                //System.Windows.Forms.MessageBox.Show("remoteObject " + (remoteObject != null).ToString());


                RemoteTestAssemblyRunner remoteTestAssemblyRunner = CreateRemoteTestAssemblyRunner(testCases, remoteObject, cancellationTokenSource, messageBus);
                //System.Windows.Forms.MessageBox.Show("remoteTestAssemblyRunner " + (remoteTestAssemblyRunner != null).ToString());
                result = remoteTestAssemblyRunner.Run();
            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }

            return result;
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