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
            string useComArgument = nameof(ExcelFactAttribute.UseCOM);

            object useCOM = GetNamedArg(testMethod.Method.GetCustomAttributes(typeof(ExcelFactAttribute)).FirstOrDefault(), useComArgument);
            if (useCOM == null)
                useCOM = GetNamedArg(testMethod.TestClass.Class.GetCustomAttributes(typeof(ExcelTestSettingsAttribute)).FirstOrDefault(), useComArgument);

            return useCOM != null ? (bool)useCOM : false;
        }

        private static object GetNamedArg(IAttributeInfo attributeInfo, string argumentName)
        {
            if (!IsNamedArg(attributeInfo, argumentName))
                return null;

            return attributeInfo.GetNamedArgument<object>(argumentName);
        }

        private static bool IsNamedArg(IAttributeInfo attributeInfo, string argumentName)
        {
            if (attributeInfo is ReflectionAttributeInfo reflectionAttributeInfo)
                return reflectionAttributeInfo.AttributeData.NamedArguments.Any(arg => arg.MemberName == argumentName);
            return false;
        }
    }
}
