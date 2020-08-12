using System;
using System.Reflection;
using Microsoft.VisualStudio.Debugger.Interop;

namespace ExcelDNAVSExtension
{
    class VSExceptionSettings
    {
        public static void DisableLoaderLock()
        {
            Type DebuggerServiceHelper = Type.GetType("VSDebugCoreUI.Utilities.DebuggerServiceHelper, VSDebugCoreUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a");
            Type ExceptionSettingsBatchOperation = Type.GetType("VSDebugCoreUI.Utilities.ExceptionSettingsBatchOperation, VSDebugCoreUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a");
            if (DebuggerServiceHelper == null || ExceptionSettingsBatchOperation == null)
                return;

            var BeginExceptionBatchOperation = DebuggerServiceHelper.GetMethod("BeginExceptionBatchOperation", BindingFlags.Public | BindingFlags.Static);
            var UpdateException = DebuggerServiceHelper.GetMethod("UpdateException", BindingFlags.Public | BindingFlags.Static);
            var EndBatchOperation = ExceptionSettingsBatchOperation.GetMethod("EndBatchOperation", BindingFlags.NonPublic | BindingFlags.Instance);
            if (BeginExceptionBatchOperation == null || UpdateException == null || EndBatchOperation == null)
                return;

            EXCEPTION_INFO150 loaderLock = new EXCEPTION_INFO150();
            loaderLock.bstrExceptionName = "LoaderLock";
            loaderLock.guidType = Guid.Parse("6ece07a9-0ede-45c4-8296-818d8fc401d4");
            loaderLock.dwState = 32770;

            object batchOperation = BeginExceptionBatchOperation.Invoke(null, null);
            UpdateException.Invoke(null, new object[] { loaderLock });
            EndBatchOperation.Invoke(batchOperation, null);
        }
    }
}
