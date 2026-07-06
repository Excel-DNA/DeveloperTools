# ExcelDna.Testing

ExcelDna.Testing helps run automated tests against Excel workbooks and Excel-DNA add-ins. Tests use xUnit and can run out-of-process by driving Excel through COM, or in-process inside Excel when access to the Excel C API is needed.

```csharp
using Xunit;

[assembly: Xunit.TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]

public class Tests
{
    [ExcelFact(OutOfProcess = true)]
    public void CanStartExcel()
    {
        Assert.Equal("16.0", ExcelDna.Testing.Util.Application.Version);
    }
}
```

Use this package for integration tests where the behavior depends on a real Excel instance, loaded add-ins, workbook state, ribbon commands, or Excel object model interaction.
