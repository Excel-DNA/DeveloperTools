# How to create Excel integration tests

## 1. Create a new .NET Framework class library C# project.

## 2. Install the ExcelDna.Testing NuGet package.

## 3. Add to the project the following Tests.cs file:

```csharp
using Xunit;

[assembly: Xunit.TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]

namespace Tests
{
    public class Tests
    {
        [ExcelFact(OutOfProcess = true)]
        public void OutOfProcessGetExcelVersion()
        {
            Assert.Equal("16.0", ExcelDna.Testing.Util.Application.Version);
        }

        [ExcelFact]
        public void InProcessGetExcelVersion()
        {
            Assert.Equal("16.0", ExcelDna.Testing.Util.Application.Version);
        }
    }
}
```

## 4. Build the project and run tests from the Test Explorer.

