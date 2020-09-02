// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelDna.Testing.VS
{
    internal class IntegrationHelper
    {
        public static EnvDTE.DTE TryLocateDteForProcess(Process process)
        {
            object dte = null;
            var monikers = new Microsoft.VisualStudio.OLE.Interop.IMoniker[1];

            NativeMethods.GetRunningObjectTable(0, out var runningObjectTable);
            runningObjectTable.EnumRunning(out var enumMoniker);
            NativeMethods.CreateBindCtx(0, out var bindContext);

            do
            {
                monikers[0] = null;
                var hresult = enumMoniker.Next(1, monikers, out _);

                if (hresult == 1)
                {
                    // There's nothing further to enumerate, so fail
                    return null;
                }
                else
                {
                    Marshal.ThrowExceptionForHR(hresult);
                }

                var moniker = monikers[0];
                moniker.GetDisplayName(bindContext, null, out var fullDisplayName);

                // FullDisplayName will look something like: <ProgID>:<ProcessId>
                var displayNameParts = fullDisplayName.Split(':');
                if (!int.TryParse(displayNameParts.Last(), out var displayNameProcessId))
                {
                    continue;
                }

                if (displayNameParts[0].StartsWith("!VisualStudio.DTE", StringComparison.OrdinalIgnoreCase) &&
                    displayNameProcessId == process.Id)
                {
                    runningObjectTable.GetObject(moniker, out dte);
                }
            }
            while (dte == null);

            return (EnvDTE.DTE)dte;
        }
    }
}
