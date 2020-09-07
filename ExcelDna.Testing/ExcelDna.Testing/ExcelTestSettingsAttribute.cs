using System;

namespace Xunit
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelTestSettingsAttribute : Attribute
    {
        public bool UseCOM { get; set; }
    }
}
