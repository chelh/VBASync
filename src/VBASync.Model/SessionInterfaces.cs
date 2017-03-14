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
        string DiffTool { get; set; }
        string DiffToolParameters { get; set; }
        string FilePath { get; set; }
        string FolderPath { get; set; }
        ISession Copy();
    }
}