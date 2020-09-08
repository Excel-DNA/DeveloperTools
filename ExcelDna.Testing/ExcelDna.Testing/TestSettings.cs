namespace ExcelDna.Testing
{
    public class TestSettings
    {
        public TestSettings(bool useCOM, string workbook)
        {
            UseCOM = useCOM;
            Workbook = workbook;
        }

        public bool UseCOM { get; }
        public string Workbook { get; }
    }
}
