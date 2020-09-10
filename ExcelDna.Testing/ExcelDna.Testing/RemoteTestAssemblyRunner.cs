using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    public sealed class RemoteTestAssemblyRunner : LongLivedMarshalByRefObject, IDisposable
    {
        private readonly XunitTestAssemblyRunner runner;

        public RemoteTestAssemblyRunner(ITestAssembly testAssembly, IEnumerable<ExcelTestCase> testCases, IMessageSink diagnosticMessageSink, IMessageSink executionMessageSink, ITestFrameworkExecutionOptions executionOptions, IMessageBus messageBus)
        {
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(delegate
            {
                ExcelDna.Testing.Util.Application = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            });
            runner = new RealTestAssemblyRunner(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions, messageBus);
        }

        public SerializableRunSummary Run()
        {
            var result = runner.RunAsync().GetAwaiter().GetResult();
            return new SerializableRunSummary
            {
                Total = result.Total,
                Failed = result.Failed,
                Skipped = result.Skipped,
                Time = result.Time,
            };
        }

        public void Dispose() => runner.Dispose();

        private class RealTestAssemblyRunner : XunitTestAssemblyRunner
        {
            private readonly IMessageBus messageBus;

            public RealTestAssemblyRunner(ITestAssembly testAssembly, IEnumerable<ExcelTestCase> testCases, IMessageSink diagnosticMessageSink, IMessageSink executionMessageSink, ITestFrameworkExecutionOptions executionOptions, IMessageBus messageBus)
                : base(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions)
            {
                this.messageBus = messageBus;
            }

            protected override IMessageBus CreateMessageBus()
            {
                return messageBus;
            }

            protected override Task<RunSummary> RunTestCollectionAsync(IMessageBus messageBus, ITestCollection testCollection, IEnumerable<IXunitTestCase> testCases, CancellationTokenSource cancellationTokenSource)
            {
                testCases = testCases.Cast<ExcelTestCase>().Select(tc => ExcelTestCase.DeserializeFromString(tc.SerializeToString()));
                return base.RunTestCollectionAsync(messageBus, testCollection, testCases, cancellationTokenSource);
            }
        }
    }
}