using ExcelDna.Integration;

namespace ExampleAddin
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        [ExcelFunction]
        public static string MyAsyncFunction()
        {
            dynamic app = ExcelDnaUtil.Application;
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            System.Threading.Tasks.Task.Factory.StartNew(() => System.Threading.Thread.Sleep(200)).ContinueWith(t =>
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    caller.SetValue("Completed");
                }));
            return "Running...";
        }
    }
}
