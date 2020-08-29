using System;
using Xunit.Sdk;
using ExcelDna.Testing;

namespace Xunit
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    [XunitTestCaseDiscoverer("ExcelDna.Testing." + nameof(ExcelFactDiscoverer), "ExcelDna.Testing")]
    public class ExcelFactAttribute : FactAttribute
    {
        /// <inheritdoc />
        public string Version { get; set; }

        /// <inheritdoc />
        public string RootSuffix { get; set; }

        /// <inheritdoc />
        public bool DebugMixedMode { get; set; }

        /// <inheritdoc />
        public bool SecureChannel { get; set; }

        /// <inheritdoc />
        public string ExtensionsDirectory { get; set; }

        /// <inheritdoc />
        public string ScreenshotsDirectory { get; set; }

        /// <inheritdoc />
        public bool UIThread { get; set; }

        /// <inheritdoc />
        public bool ReuseInstance { get; set; }

        /// <inheritdoc />
        public bool TakeScreenshotOnFailure { get; set; }
    }
}
