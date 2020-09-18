using ExcelDna.Integration;

namespace ExcelAgent
{
    /// <summary>
    /// This command is required to properly initialize the agent in Excel.
    /// </summary>
    public static class InitializationCommand
    {
        [ExcelCommand(MenuName = "ExcelAgent", MenuText = "Internal initialization")]
        public static void InternalInitialization()
        {
        }
    }
}
