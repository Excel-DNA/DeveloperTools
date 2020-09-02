// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System.Runtime.InteropServices;

namespace ExcelDna.Testing.VS
{
    internal class NativeMethods
    {
        private const string Ole32 = "ole32.dll";

        [DllImport(Ole32, PreserveSig = false)]
        public static extern void CreateBindCtx(int reserved, [MarshalAs(UnmanagedType.Interface)] out Microsoft.VisualStudio.OLE.Interop.IBindCtx bindContext);

        [DllImport(Ole32, PreserveSig = false)]
        public static extern void GetRunningObjectTable(int reserved, [MarshalAs(UnmanagedType.Interface)] out Microsoft.VisualStudio.OLE.Interop.IRunningObjectTable runningObjectTable);
    }
}
