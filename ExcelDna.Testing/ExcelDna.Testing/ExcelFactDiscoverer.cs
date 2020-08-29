using System.Collections.Generic;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    class ExcelFactDiscoverer : FactDiscoverer
    {
        public ExcelFactDiscoverer(IMessageSink diagnosticMessageSink)
            : base(diagnosticMessageSink)
        {
        }

        public override IEnumerable<IXunitTestCase> Discover(ITestFrameworkDiscoveryOptions discoveryOptions, ITestMethod testMethod, IAttributeInfo factAttribute)
        {
            //System.Windows.Forms.MessageBox.Show(testMethod.Method.Name);
            var results = new List<IXunitTestCase>();
            results.Add(new ExcelTestCase(DiagnosticMessageSink, discoveryOptions.MethodDisplayOrDefault(), testMethod, null));
            return results;
        }
    }
}
