namespace VbaSync {
    public class MainViewModel : ViewModelBase, ISession {
        ActionType _action = ActionType.Extract;
        string _diffTool;
        string _diffToolParameters;
        string _filePath;
        string _folderPath;

        public ActionType Action {
            get { return _action; }
            set { SetField(ref _action, value, nameof(Action)); }
        }

        public string DiffTool {
            get { return _diffTool; }
            set { SetField(ref _diffTool, value, nameof(DiffTool)); }
        }

        public string DiffToolParameters {
            get { return _diffToolParameters; }
            set { SetField(ref _diffToolParameters, value, nameof(DiffToolParameters)); }
        }

        public string FilePath {
            get { return _filePath; }
            set { SetField(ref _filePath, value, nameof(FilePath)); }
        }

        public string FolderPath {
            get { return _folderPath; }
            set { SetField(ref _folderPath, value, nameof(FolderPath)); }
        }

        public ISession Copy() => new MainViewModel {
            _action = _action,
            _diffTool = _diffTool,
            _diffToolParameters = _diffToolParameters,
            _filePath = _filePath,
            _folderPath = _folderPath
        };
    }
}
