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

            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = excelExePath;
            info.Arguments = "\"" + GetXllPath(addinAssemblyPath, bitness) + "\"";
            return Process.Start(info);
        }

        public static string GetXllPath(string addinAssemblyPath, Bitness bitness)
        {
            return Path.Combine(Path.GetDirectoryName(addinAssemblyPath), Path.GetFileNameWithoutExtension(addinAssemblyPath) + "-AddIn" + (bitness == Bitness.Bit64 ? "64" : "") + ".xll");
        }

        private string excelExePath;
        private Bitness bitness;
        private bool excelDetected;
    }
}
