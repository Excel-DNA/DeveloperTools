using Newtonsoft.Json;
using System;

namespace ExcelDna.Testing.Remote
{
    public class ExcelTestCaseConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            ExcelTestCase test = (ExcelTestCase)value;

            writer.WriteValue(test.SerializeToString());
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (reader.Value == null)
                return null;

            return ExcelTestCase.DeserializeFromString((string)reader.Value);
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(ExcelTestCase);
        }
    }

    public class TestCaseConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (reader.Value == null)
                return null;

            return ExcelTestCase.DeserializeFromString((string)reader.Value);
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(Xunit.Abstractions.ITestCase) || objectType == typeof(Xunit.Sdk.IXunitTestCase);
        }
    }

    public class ReflectionAssemblyInfoConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            Xunit.Sdk.ReflectionAssemblyInfo info = (Xunit.Sdk.ReflectionAssemblyInfo)value;

            writer.WriteValue(info.AssemblyPath);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (reader.Value == null)
                return null;

            return new Xunit.Sdk.ReflectionAssemblyInfo((string)reader.Value);
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(Xunit.Abstractions.IAssemblyInfo) || objectType == typeof(Xunit.Sdk.ReflectionAssemblyInfo);
        }
    }

    public class ReflectionTypeInfoConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            Xunit.Sdk.ReflectionTypeInfo info = (Xunit.Sdk.ReflectionTypeInfo)value;

            writer.WriteValue(info.Type.FullName);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (reader.Value == null)
                return null;

            return new Xunit.Sdk.ReflectionTypeInfo(Type.GetType((string)reader.Value));
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(Xunit.Sdk.ReflectionTypeInfo) || objectType == typeof(Xunit.Abstractions.ITypeInfo);
        }
    }
}
