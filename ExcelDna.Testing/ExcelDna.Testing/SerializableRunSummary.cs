using System;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    [Serializable]
    public class SerializableRunSummary
    {
        public int Total { get; set; }
        public int Failed { get; set; }
        public int Skipped { get; set; }
        public decimal Time { get; set; }

        public static implicit operator RunSummary(SerializableRunSummary summary) => new RunSummary
        {
            Total = summary.Total,
            Failed = summary.Failed,
            Skipped = summary.Skipped,
            Time = summary.Time,
        };
    }
}