namespace ExcelDna.Testing
{
    public class ExcelTestSettings : ITestSettings
    {
        public ExcelTestSettings(bool outOfProcess, string workbook, string addin)
        {
            OutOfProcess = outOfProcess;
            Workbook = workbook;
            AddIn = addin;
        }

        public bool OutOfProcess { get; set; }

        public string Workbook { get; set; }

        public string AddIn { get; set; }
    }
}
