using VBASync.Model;

namespace VBASync.Tests.Mocks
{
    public class QuickSessionSettings : ISessionSettings
    {
        public bool AddNewDocumentsToFile { get; set; }
        public Hook AfterExtractHook { get; set; }
        public Hook BeforePublishHook { get; set; }
        public bool DeleteDocumentsFromFile { get; set; }
        public bool IgnoreEmpty { get; set; }
        public bool SearchRepositorySubdirectories { get; set; }
    }
}
