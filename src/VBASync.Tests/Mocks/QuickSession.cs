using VBASync.Model;

namespace VBASync.Tests.Mocks
{
    internal class QuickSession : ISession
    {
        public ActionType Action { get; set; }
        public bool AutoRun { get; set; }
        public string FilePath { get; set; }
        public string FolderPath { get; set; }
    }
}
