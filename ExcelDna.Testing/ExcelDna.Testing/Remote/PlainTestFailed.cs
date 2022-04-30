using Xunit.Abstractions;

namespace ExcelDna.Testing.Remote
{
    internal class PlainTestFailed
    {
        public decimal ExecutionTime { get; set; }

        public string Output { get; set; }

        public ITest Test { get; set; }

        public string[] ExceptionTypes { get; set; }

        public string[] Messages { get; set; }

        public string[] StackTraces { get; set; }

        public int[] ExceptionParentIndices { get; set; }
    }
}
