using ExcelDna.Integration;

namespace ExampleAddinNET6
{
    public class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET 6 function")]
        public static string SayHelloNET6(string name)
        {
            return "Hello .NET6 " + name;
        }
    }
}