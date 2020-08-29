using System.Reflection;
using Xunit.Abstractions;
using Xunit.Sdk;
using ExcelDna.Testing;



namespace Xunit
{
    internal class ExcelTestFramework : XunitTestFramework
    {
        public ExcelTestFramework(IMessageSink messageSink) : base(messageSink)
        {
        }

        protected override ITestFrameworkExecutor CreateExecutor(AssemblyName assemblyName)
        {
            return new ExcelTestFrameworkExecutor(assemblyName, SourceInformationProvider, DiagnosticMessageSink);
        }
    }
}
