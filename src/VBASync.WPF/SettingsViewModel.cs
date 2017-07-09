namespace VBASync.WPF
{
    internal class SettingsViewModel : ViewModelBase, Model.ISessionSettings
    {
        private bool _addNewDocumentsToFile;
        private string _afterExtractHookContent;
        private string _beforePublishHookContent;
        private bool _deleteDocumentsFromFile;
        private string _diffTool;
        private string _diffToolParameters;
        private bool _ignoreEmpty;
        private string _language;
        private bool _portable;
        private bool _searchRepositorySubdirectories;

        public bool AddNewDocumentsToFile
        {
            get => _addNewDocumentsToFile;
            set => SetField(ref _addNewDocumentsToFile, value, nameof(AddNewDocumentsToFile));
        }

        public Model.Hook AfterExtractHook => new Model.Hook(AfterExtractHookContent);

        public string AfterExtractHookContent
        {
            get => _afterExtractHookContent;
            set => SetField(ref _afterExtractHookContent, value, nameof(AfterExtractHookContent));
        }

        public Model.Hook BeforePublishHook => new Model.Hook(BeforePublishHookContent);

        public string BeforePublishHookContent
        {
            get => _beforePublishHookContent;
            set => SetField(ref _beforePublishHookContent, value, nameof(BeforePublishHook));
        }

        public bool DeleteDocumentsFromFile
        {
            get => _deleteDocumentsFromFile;
            set => SetField(ref _deleteDocumentsFromFile, value, nameof(DeleteDocumentsFromFile));
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

        public bool IgnoreEmpty
        {
            get => _ignoreEmpty;
            set => SetField(ref _ignoreEmpty, value, nameof(IgnoreEmpty));
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

        public bool SearchRepositorySubdirectories
        {
            get => _searchRepositorySubdirectories;
            set => SetField(ref _searchRepositorySubdirectories, value, nameof(SearchRepositorySubdirectories));
        }

        public SettingsViewModel Clone() => new SettingsViewModel
        {
            _addNewDocumentsToFile = _addNewDocumentsToFile,
            _afterExtractHookContent = _afterExtractHookContent,
            _beforePublishHookContent = _beforePublishHookContent,
            _deleteDocumentsFromFile = _deleteDocumentsFromFile,
            _diffTool = _diffTool,
            _diffToolParameters = _diffToolParameters,
            _ignoreEmpty = _ignoreEmpty,
            _language = _language,
            _portable = _portable,
            _searchRepositorySubdirectories = _searchRepositorySubdirectories
        };
    }
}
