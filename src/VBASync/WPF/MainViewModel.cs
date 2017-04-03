namespace VBASync.WPF
{
    internal class MainViewModel : ViewModelBase
    {
        private SessionViewModel _session;
        private SettingsViewModel _settings;

        public SessionViewModel Session
        {
            get => _session;
            set => SetField(ref _session, value, nameof(Session));
        }

        public SettingsViewModel Settings
        {
            get => _settings;
            set => SetField(ref _settings, value, nameof(Settings));
        }
    }
}
