namespace ExcelDna.Testing
{
    public static class ExcelStartupEvent
    {
        public static void Create()
        {
            if (eventHandle == null)
                eventHandle = new System.Threading.EventWaitHandle(false, System.Threading.EventResetMode.ManualReset, eventName);

            eventHandle.Reset();
        }

        public static void Set()
        {
            System.Threading.EventWaitHandle.OpenExisting(eventName).Set();
        }

        public static bool Wait(int millisecondsTimeout)
        {
            return eventHandle.WaitOne(millisecondsTimeout);
        }

        private static System.Threading.EventWaitHandle eventHandle;
        private const string eventName = "ExcelDna.Testing.ExcelStartupEvent.e275d37a-c1cc-435e-9b51-96f7648717c1";
    }
}
