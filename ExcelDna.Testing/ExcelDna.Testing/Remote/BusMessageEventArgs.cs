using System;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing.Remote
{
    public class BusMessageEventArgs : EventArgs
    {
        public BusMessageEventArgs()
        {
        }

        public BusMessageEventArgs(IMessageSinkMessage message)
        {
            if (message is TestMethodMessage tmm)
            {
                if (tmm.TestMethod is TestMethod tm)
                {
                    if (tm.Method is ReflectionMethodInfo)
                    {
                        tm.Method = new PlainMethodInfo(tm.Method);
                    }
                }
            }

            try
            {
                SerializedMessage = Newtonsoft.Json.JsonConvert.SerializeObject(message, serializerSettings);
            }
            catch
            {
            }
        }

        public IMessageSinkMessage GetMessage()
        {
            if (SerializedMessage == null)
                return null;

            try
            {
                if (SerializedMessage.Contains(testFailed))
                {
                    SerializedMessage = SerializedMessage.Replace(testFailed, "ExcelDna.Testing.Remote.PlainTestFailed, ExcelDna.Testing");
                    var plain = Newtonsoft.Json.JsonConvert.DeserializeObject<PlainTestFailed>(SerializedMessage, deserializerSettings);
                    return new TestFailed(plain.Test, plain.ExecutionTime, plain.Output, plain.ExceptionTypes, plain.Messages, plain.StackTraces, plain.ExceptionParentIndices);
                }

                return Newtonsoft.Json.JsonConvert.DeserializeObject<IMessageSinkMessage>(SerializedMessage, deserializerSettings);
            }
            catch
            {
            }

            return null;
        }

        public string SerializedMessage { get; set; }

        private static Newtonsoft.Json.JsonSerializerSettings serializerSettings = new Newtonsoft.Json.JsonSerializerSettings()
        {
            TypeNameHandling = Newtonsoft.Json.TypeNameHandling.All,
            Converters = new Newtonsoft.Json.JsonConverter[] { new ReflectionAssemblyInfoConverter(), new ReflectionTypeInfoConverter() }
        };

        private static Newtonsoft.Json.JsonSerializerSettings deserializerSettings = new Newtonsoft.Json.JsonSerializerSettings()
        {
            TypeNameHandling = Newtonsoft.Json.TypeNameHandling.All,
            Converters = new Newtonsoft.Json.JsonConverter[] { new TestCaseConverter(), new ReflectionAssemblyInfoConverter(), new ReflectionTypeInfoConverter() }
        };

        private static string testFailed = "Xunit.Sdk.TestFailed, xunit.execution.desktop";
    }
}
