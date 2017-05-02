namespace VBASync.Model
{
    public enum ActionType
    {
        Extract = 0,
        Publish = 1
    }

    public interface ISession
    {
        ActionType Action { get; set; }
        bool AutoRun { get; set; }
        string FilePath { get; set; }
        string FolderPath { get; set; }
    }

    public interface ISessionSettings
    {
        bool AddNewDocumentsToFile { get; set; }
        bool IgnoreEmpty { get; set; }
    }
}
