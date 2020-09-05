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

        public ExcelTestCase(bool useCOM, IMessageSink diagnosticMessageSink, TestMethodDisplay defaultMethodDisplay, ITestMethod testMethod, object[] testMethodArguments = null)
            : base(diagnosticMessageSink, defaultMethodDisplay, testMethod, testMethodArguments)
        {
            this.useCOM = useCOM;
        }

        public bool UseCOM => useCOM;

        public override Task<RunSummary> RunAsync(IMessageSink diagnosticMessageSink, IMessageBus messageBus, object[] constructorArguments, ExceptionAggregator aggregator, CancellationTokenSource cancellationTokenSource)
    => new ExcelTestCaseRunner(this, DisplayName, SkipReason, constructorArguments, TestMethodArguments, messageBus, aggregator, cancellationTokenSource).RunAsync();

        public override void Serialize(IXunitSerializationInfo info)
        {
            base.Serialize(info);
            info.AddValue(nameof(useCOM), useCOM);
        }

        public override void Deserialize(IXunitSerializationInfo info)
        {
            base.Deserialize(info);
            useCOM = info.GetValue<bool>(nameof(useCOM));
        }

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

        private bool useCOM;
    }
}
