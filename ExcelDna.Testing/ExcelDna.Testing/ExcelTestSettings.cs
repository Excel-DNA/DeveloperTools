namespace ExcelDna.Testing
{
    public class ExcelTestSettings : ITestSettings
    {
        public ExcelTestSettings(bool outOfProcess, string workbook)
        {
            OutOfProcess = outOfProcess;
            Workbook = workbook;
        }

        public bool OutOfProcess { get; set; }

        public string Workbook { get; set; }
    }
}
