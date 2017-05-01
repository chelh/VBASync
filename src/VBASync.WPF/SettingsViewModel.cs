namespace VBASync.WPF
{
    internal class SettingsViewModel : ViewModelBase, Model.ISessionSettings
    {
        private bool _addNewDocumentsToFile;
        private string _diffTool;
        private string _diffToolParameters;
        private string _language;
        private bool _portable;

        public bool AddNewDocumentsToFile
        {
            get => _addNewDocumentsToFile;
            set => SetField(ref _addNewDocumentsToFile, value, nameof(AddNewDocumentsToFile));
        }

        public string DiffTool
        {
            get => _diffTool;
            set => SetField(ref _diffTool, value, nameof(DiffTool));
        }

        public string DiffToolParameters
        {
            get => _diffToolParameters;
            set => SetField(ref _diffToolParameters, value, nameof(DiffToolParameters));
        }

        public string Language
        {
            get => _language;
            set => SetField(ref _language, value, nameof(Language));
        }

        public bool Portable
        {
            get => _portable;
            set => SetField(ref _portable, value, nameof(Portable));
        }

        public SettingsViewModel Clone() => new SettingsViewModel
        {
            _addNewDocumentsToFile = _addNewDocumentsToFile,
            _diffTool = _diffTool,
            _diffToolParameters = _diffToolParameters,
            _language = _language,
            _portable = _portable
        };
    }
}
