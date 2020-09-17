namespace ExcelDna.Testing
{
    public class ExcelTestSettings : ITestSettings
    {
        public ExcelTestSettings(bool outOfProcess, string workbook, string xll)
        {
            OutOfProcess = outOfProcess;
            Workbook = workbook;
            XLL = xll;
        }

        public bool OutOfProcess { get; set; }

        public string Workbook { get; set; }

        public string XLL { get; set; }
    }
}
