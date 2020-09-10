using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    internal class ExcelTestRunner : XunitTestRunner
    {
        public ExcelTestRunner(ITest test, IMessageBus messageBus, Type testClass, object[] constructorArguments, MethodInfo testMethod, object[] testMethodArguments, string skipReason, IReadOnlyList<BeforeAfterTestAttribute> beforeAfterAttributes, ExceptionAggregator aggregator, CancellationTokenSource cancellationTokenSource)
            : base(test, messageBus, testClass, constructorArguments, testMethod, testMethodArguments, skipReason, beforeAfterAttributes, aggregator, cancellationTokenSource)
        {
        }

        protected new ExcelTestCase TestCase => (ExcelTestCase)base.TestCase;

        protected override Task<decimal> InvokeTestMethodAsync(ExceptionAggregator aggregator)
        {
            var result = !TestCase.Settings.OutOfProcess
               ? InvokeOnUIThreadAsync(aggregator)
               : InvokeAsync(aggregator);

            return result;
        }

        private Task<decimal> InvokeOnUIThreadAsync(ExceptionAggregator aggregator)
        {
            var tcs = new TaskCompletionSource<decimal>();

            ExcelAction a = async delegate
            {
                try
                {
                    var result = await InvokeAsync(aggregator);
                    tcs.SetResult(result);
                }
                catch (Exception e)
                {
                    tcs.SetException(e);
                }
            };
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(a);

            return tcs.Task;
        }

        private Task<decimal> InvokeAsync(ExceptionAggregator aggregator)
        {
            if (TestCase.Settings.Workbook != null)
            {
                if (TestCase.Settings.Workbook.Length == 0)
                {
                    Util.Workbook = Util.Application.Workbooks.Add();
                }
                else
                {
                    string workbookPath = Path.Combine(Util.TestAssemblyDirectory, TestCase.Settings.Workbook);
                    Util.Workbook = Util.Application.Workbooks.Open(workbookPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                }
            }

            try
            {
                return new XunitTestInvoker(Test, MessageBus, TestClass, ConstructorArguments, TestMethod, TestMethodArguments, BeforeAfterAttributes, aggregator, CancellationTokenSource).RunAsync();
            }
            finally
            {
                if (Util.Workbook != null)
                {
                    Util.Workbook.Close(false);
                    Util.Workbook = null;
                }
            }
        }
    }
}
