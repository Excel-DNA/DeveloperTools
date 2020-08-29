using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace Examples
{
    class Utils
    {
        public static string GetVersion()
        {
            Application app = (Application)ExcelDnaUtil.Application;
            return app.Version;
            //return "xxx";
        }
    }
}
