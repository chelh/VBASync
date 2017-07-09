namespace VBASync.Model
{
    public enum ActionType
    {
        Extract = 0,
        Publish = 1
    }

    public interface ISession
    {
        ActionType Action { get; }
        bool AutoRun { get; }
        string FilePath { get; }
        string FolderPath { get; }
    }

    public interface ISessionSettings
    {
        bool AddNewDocumentsToFile { get; }
        Hook AfterExtractHook { get; }
        Hook BeforePublishHook { get; }
        bool DeleteDocumentsFromFile { get; }
        bool IgnoreEmpty { get; }
        bool SearchRepositorySubdirectories { get; }
    }
}
