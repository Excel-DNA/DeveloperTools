using Microsoft.Office.Interop.Excel;
using Xunit;

namespace Examples
{
    [ExcelTestSettings(UseCOM = true)]
    public class ComsTest
    {
        [ExcelFact]
        public void COMGetExcelVersion()
        {
            Assert.Equal("16.0", ExcelDna.Testing.COMUtil.Application.Version);
        }

        [ExcelFact(Workbook = "")]
        public void COMFunctionSayHello()
        {
            Range functionRange = ExcelDna.Testing.COMUtil.Workbook.Sheets[1].Range["B1:B1"];
            functionRange.Formula = "=SayHello(\"world\")";
            Assert.Equal("Hello world", functionRange.Value.ToString());
        }
    }
}
