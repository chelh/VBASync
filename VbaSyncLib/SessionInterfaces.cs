using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaSync {
    public enum ActionType {
        Extract,
        Publish
    }

    public interface ISession {
        ActionType Action { get; set; }
        string DiffTool { get; set; }
        string DiffToolParameters { get; set; }
        string FilePath { get; set; }
        string FolderPath { get; set; }
        ISession Copy();
    }
}