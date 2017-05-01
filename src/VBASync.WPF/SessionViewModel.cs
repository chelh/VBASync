namespace VBASync.WPF
{
    internal class SessionViewModel : ViewModelBase, Model.ISession
    {
        private Model.ActionType _action;
        private bool _autoRun;
        private string _filePath;
        private string _folderPath;

        public Model.ActionType Action
        {
            get { return _action; }
            set { SetField(ref _action, value, nameof(Action)); }
        }

        public bool AutoRun
        {
            get { return _autoRun; }
            set { SetField(ref _autoRun, value, nameof(AutoRun)); }
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

        public Model.ISession Copy() => new SessionViewModel
        {
            _action = _action,
            _autoRun = _autoRun,
            _filePath = _filePath,
            _folderPath = _folderPath
        };
    }
}
