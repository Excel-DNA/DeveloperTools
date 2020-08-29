using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    public class RemoteObject : LongLivedMarshalByRefObject
    {
        public RemoteTestAssemblyRunner CreateTestAssemblyRunner(string testAssemblyPath, string testAssemblyConfigurationFile, ExcelTestCase[] testCases, IMessageSink diagnosticMessageSink, IMessageSink executionMessageSink, ITestFrameworkExecutionOptions executionOptions, IMessageBus messageBus)
        {
            var testAssembly = new TestAssembly(new ReflectionAssemblyInfo(Assembly.LoadFrom(testAssemblyPath)), testAssemblyConfigurationFile);
            return new RemoteTestAssemblyRunner(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions, messageBus);
        }
    }
}
