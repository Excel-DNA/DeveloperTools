namespace ExcelDna.Testing
{
    public static class COMUtil
    {
        public static Microsoft.Office.Interop.Excel.Application Application => application;

        internal static void SetApplication(Microsoft.Office.Interop.Excel.Application newApplication)
        {
            application = newApplication;
        }

        private static Microsoft.Office.Interop.Excel.Application application;
    }
}
