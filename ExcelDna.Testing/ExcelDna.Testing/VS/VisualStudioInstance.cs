// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System.Diagnostics;
using System.Linq;

namespace ExcelDna.Testing.VS
{
    internal class VisualStudioInstance
    {
        public static void AttachDebugger(Process hostProcess)
        {
            var debuggerHostDte = GetDebuggerHostDte();
            var targetProcessId = Process.GetCurrentProcess().Id;
            var localProcess = debuggerHostDte?.Debugger.LocalProcesses.OfType<EnvDTE80.Process2>().FirstOrDefault(p => p.ProcessID == hostProcess.Id);
            if (localProcess != null)
            {
                localProcess.Attach2("Managed");
            }
        }

        private static EnvDTE.DTE GetDebuggerHostDte()
        {
            var currentProcessId = Process.GetCurrentProcess().Id;
            foreach (var process in Process.GetProcessesByName("devenv"))
            {
                var dte = IntegrationHelper.TryLocateDteForProcess(process);
                if (dte?.Debugger?.DebuggedProcesses?.OfType<EnvDTE.Process>().Any(p => p.ProcessID == currentProcessId) ?? false)
                {
                    return dte;
                }
            }

            return null;
        }
    }
}
