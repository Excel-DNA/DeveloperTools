using ExcelDna.Integration;
using System;
using System.Collections.Generic;
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
            var result = !TestCase.Settings.UseCOM
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
            => new XunitTestInvoker(Test, MessageBus, TestClass, ConstructorArguments, TestMethod, TestMethodArguments, BeforeAfterAttributes, aggregator, CancellationTokenSource).RunAsync();
    }
}
