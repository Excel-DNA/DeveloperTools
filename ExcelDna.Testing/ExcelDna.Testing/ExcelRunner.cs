using System;
using System.Diagnostics;
using System.IO;

namespace ExcelDna.Testing
{
    internal class ExcelRunner
    {
        public ExcelRunner()
        {
            ExcelDetector excelDetector = new ExcelDetector();
            excelDetected = excelDetector.TryFindLatestExcel(out excelExePath) && excelDetector.TryFindExcelBitness(excelExePath, out bitness);
        }

        public Process Start(string addinAssemblyPath)
        {
            if (!excelDetected)
                throw new ApplicationException("Can't find an installed version of Excel.");

            string externalXllRelativePath = @"..\..\..\ExampleAddin\bin\Debug\ExampleAddin-AddIn";
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = excelExePath;
            info.Arguments = Quote(GetXllPath(addinAssemblyPath, externalXllRelativePath, bitness)) + " " + Quote(GetTestsXllPath(addinAssemblyPath, bitness));
            return Process.Start(info);
        }

        public static string GetTestsXllPath(string addinAssemblyPath, Bitness bitness)
        {
            return GetXllPath(addinAssemblyPath, Path.GetFileNameWithoutExtension(addinAssemblyPath) + "-AddIn", bitness);
        }

        public static string GetXllPath(string addinAssemblyPath, string externalXllRelativePath, Bitness bitness)
        {
            return Path.Combine(Path.GetDirectoryName(addinAssemblyPath), externalXllRelativePath + (bitness == Bitness.Bit64 ? "64" : "") + ".xll");
        }

        private string Quote(string s)
        {
            return "\"" + s + "\"";
        }

        private string excelExePath;
        private Bitness bitness;
        private bool excelDetected;
    }
}
