using System;
using System.Threading;
using System.Threading.Tasks;
using Xunit.Abstractions;
using Xunit.Internal;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    public class ExcelTestCase : XunitTestCase
    {
        [Obsolete("Called by the de-serializer; should only be called by deriving classes for de-serialization purposes")]
        public ExcelTestCase()
        {
        }

        public ExcelTestCase(IMessageSink diagnosticMessageSink, TestMethodDisplay defaultMethodDisplay, ITestMethod testMethod, object[] testMethodArguments = null)
            : base(diagnosticMessageSink, defaultMethodDisplay, testMethod, testMethodArguments)
        {
        }

        public override Task<RunSummary> RunAsync(IMessageSink diagnosticMessageSink, IMessageBus messageBus, object[] constructorArguments, ExceptionAggregator aggregator, CancellationTokenSource cancellationTokenSource)
    => new ExcelTestCaseRunner(this, DisplayName, SkipReason, constructorArguments, TestMethodArguments, messageBus, aggregator, cancellationTokenSource).RunAsync();


        public string SerializeToString()
        {
            var triple = new XunitSerializationTriple(nameof(ExcelTestCase), this, GetType());
            return XunitSerializationInfo.SerializeTriple(triple);
        }

        public static ExcelTestCase DeserializeFromString(string value)
        {
            var triple = XunitSerializationInfo.DeserializeTriple(value);
            return (ExcelTestCase)triple.Value;
        }
    }
}
