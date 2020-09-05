using System.Collections.Generic;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;
using System.Linq;

namespace ExcelDna.Testing
{
    class ExcelFactDiscoverer : FactDiscoverer
    {
        public ExcelFactDiscoverer(IMessageSink diagnosticMessageSink) : base(diagnosticMessageSink)
        {
        }

        public override IEnumerable<IXunitTestCase> Discover(ITestFrameworkDiscoveryOptions discoveryOptions, ITestMethod testMethod, IAttributeInfo factAttribute)
        {
            var results = new List<IXunitTestCase>();
            results.Add(new ExcelTestCase(UsingCOM(testMethod), DiagnosticMessageSink, discoveryOptions.MethodDisplayOrDefault(), testMethod, null));
            return results;
        }

        private static bool UsingCOM(ITestMethod testMethod)
        {
            var methodAttr = testMethod.Method.GetCustomAttributes(typeof(ExcelFactAttribute)).FirstOrDefault();
            var useCOM = methodAttr?.GetNamedArgument<object>(nameof(ExcelFactAttribute.UseCOM));
            return useCOM != null && (bool)useCOM;
        }
    }
}
