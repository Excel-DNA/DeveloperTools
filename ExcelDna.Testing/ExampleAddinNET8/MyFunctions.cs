using ExcelDna.Integration;

namespace ExampleAddinNET8
{
    public class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET 8 function")]
        public static string SayHelloNET8(string name)
        {
            return "Hello .NET8 " + name;
        }
    }
}