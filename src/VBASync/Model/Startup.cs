using System.Collections.Generic;

namespace VBASync.Model
{
    public class Startup : ISession, ISessionSettings
    {
        public ActionType Action { get; set; }
        public bool AddNewDocumentsToFile { get; set; }
        public bool AutoRun { get; set; }
        public string DiffTool { get; set; }
        public string DiffToolParameters { get; set; }
        public string FilePath { get; set; }
        public string FolderPath { get; set; }
        public string Language { get; set; }
        public string LastSessionPath { get; set; }
        public bool Portable { get; set; }
        public List<string> RecentFiles { get; } = new List<string>();
    }
}
