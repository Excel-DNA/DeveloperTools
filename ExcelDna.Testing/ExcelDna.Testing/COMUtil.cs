namespace ExcelDna.Testing
{
    public static class COMUtil
    {
        public static Microsoft.Office.Interop.Excel.Application Application => application;
        public static Microsoft.Office.Interop.Excel.Workbook Workbook => workbook;
        public static string TestAssemblyDirectory => testAssemblyDirectory;

        internal static void SetApplication(Microsoft.Office.Interop.Excel.Application newApplication)
        {
            application = newApplication;
        }

        internal static void SetWorkbook(Microsoft.Office.Interop.Excel.Workbook newWorkbook)
        {
            workbook = newWorkbook;
        }

        internal static void SetTestAssemblyDirectory(string directory)
        {
            testAssemblyDirectory = directory;
        }

        private static Microsoft.Office.Interop.Excel.Application application;
        private static Microsoft.Office.Interop.Excel.Workbook workbook;
        private static string testAssemblyDirectory;
    }
}
