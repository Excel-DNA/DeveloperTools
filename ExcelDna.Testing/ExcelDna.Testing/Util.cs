namespace ExcelDna.Testing
{
    public static class Util
    {
        public static Microsoft.Office.Interop.Excel.Application Application { get; internal set; }
        public static Microsoft.Office.Interop.Excel.Workbook Workbook { get; internal set; }
        public static string TestAssemblyDirectory { get; internal set; }
    }
}
