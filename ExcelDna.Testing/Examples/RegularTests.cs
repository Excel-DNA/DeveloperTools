using Xunit;

[assembly: Xunit.TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]

namespace Examples
{
    public class RegularTests
    {
        [Fact]
        public void RegularTest()
        {
            Assert.Equal(4, 2 + 2);
        }
    }
}
