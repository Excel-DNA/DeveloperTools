using System.Collections.Generic;
using System.IO;
using Xunit.Abstractions;

namespace ExcelDna.Testing.Remote
{
    public class AssemblyInfo : IAssemblyInfo
    {
        public AssemblyInfo(string path)
        {
            AssemblyPath = path;
        }

        public string AssemblyPath { get; }

        public string Name => Path.GetFileNameWithoutExtension(AssemblyPath);

        public IEnumerable<IAttributeInfo> GetCustomAttributes(string assemblyQualifiedAttributeTypeName)
        {
            throw new System.NotImplementedException();
        }

        public ITypeInfo GetType(string typeName)
        {
            throw new System.NotImplementedException();
        }

        public IEnumerable<ITypeInfo> GetTypes(bool includePrivateTypes)
        {
            throw new System.NotImplementedException();
        }
    }
}
