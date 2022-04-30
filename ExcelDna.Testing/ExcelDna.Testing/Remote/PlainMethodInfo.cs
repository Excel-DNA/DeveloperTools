using System.Collections.Generic;
using Xunit.Abstractions;

namespace ExcelDna.Testing.Remote
{
    public class PlainMethodInfo : IMethodInfo
    {
        public PlainMethodInfo()
        {
        }

        public PlainMethodInfo(IMethodInfo src)
        {
            Name = src.Name;
        }

        public bool IsAbstract { get; set; }

        public bool IsGenericMethodDefinition { get; set; }

        public bool IsPublic { get; set; }

        public bool IsStatic { get; set; }

        public string Name { get; set; }

        public ITypeInfo ReturnType => null;

        public ITypeInfo Type => null;

        public IEnumerable<IAttributeInfo> GetCustomAttributes(string assemblyQualifiedAttributeTypeName)
        {
            return null;
        }

        public IEnumerable<ITypeInfo> GetGenericArguments()
        {
            return null;
        }

        public IEnumerable<IParameterInfo> GetParameters()
        {
            return null;
        }

        public IMethodInfo MakeGenericMethod(params ITypeInfo[] typeArguments)
        {
            return null;
        }
    }
}
