using Microsoft.Office.Interop.Excel;
using Xunit;

namespace Examples
{
    [ExcelTestSettings(UseCOM = true)]
    public class ComTests
    {
        [ExcelFact]
        public void GetExcelVersion()
        {
            Assert.Equal("16.0", ExcelDna.Testing.Util.Application.Version);
        }

        [ExcelFact(Workbook = "")]
        public void FunctionSayHello()
        {
            Range functionRange = ExcelDna.Testing.Util.Workbook.Sheets[1].Range["B1:B1"];
            functionRange.Formula = "=SayHello(\"world\")";
            Assert.Equal("Hello world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "MrExcel.xlsx")]
        public void PreCreatedWorkbook()
        {
            Range cell = ExcelDna.Testing.Util.Workbook.Sheets[1].Range["A2:A2"];
            Assert.Equal("Red Ford Truck", cell.Value.ToString());
        }
    }
}
