using VBASync.Model;

namespace VBASync.WPF
{
    public class MainViewModel : ViewModelBase, ISession
    {
        private ActionType _action;
        private bool _autoRun;
        private string _diffTool;
        private string _diffToolParameters;
        private string _filePath;
        private string _folderPath;
        private string _language;

        public ActionType Action
        {
            get { return _action; }
            set { SetField(ref _action, value, nameof(Action)); }
        }

        public bool AutoRun
        {
            get { return _autoRun; }
            set { SetField(ref _autoRun, value, nameof(AutoRun)); }
        }

        public string DiffTool
        {
            get { return _diffTool; }
            set { SetField(ref _diffTool, value, nameof(DiffTool)); }
        }

        public string DiffToolParameters
        {
            get { return _diffToolParameters; }
            set { SetField(ref _diffToolParameters, value, nameof(DiffToolParameters)); }
        }

        public string FilePath
        {
            get { return _filePath; }
            set { SetField(ref _filePath, value, nameof(FilePath)); }
        }

        public string FolderPath
        {
            get { return _folderPath; }
            set { SetField(ref _folderPath, value, nameof(FolderPath)); }
        }

        public string Language
        {
            get { return _language; }
            set { SetField(ref _language, value, nameof(Language)); }
        }

        public ISession Copy() => new MainViewModel {
            _action = _action,
            _autoRun = _autoRun,
            _diffTool = _diffTool,
            _diffToolParameters = _diffToolParameters,
            _filePath = _filePath,
            _folderPath = _folderPath,
            _language = _language
        };
    }
}
