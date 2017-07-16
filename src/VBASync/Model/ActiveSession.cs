using System;
using System.Collections.Generic;
using System.Linq;

namespace VBASync.Model
{
    public sealed class ActiveSession : IDisposable
    {
        private readonly VbaRepositoryFolder _repositoryFolder;
        private readonly ISession _session;
        private readonly ISessionSettings _sessionSettings;
        private readonly ISystemOperations _so;
        private readonly VbaTemporaryFolder _tempFolder;

        public ActiveSession(ISession session, ISessionSettings sessionSettings) : this(new RealSystemOperations(), session, sessionSettings)
        {
        }

        internal ActiveSession(ISystemOperations so, ISession session, ISessionSettings sessionSettings)
        {
            _so = so;
            _session = session;
            _sessionSettings = sessionSettings;
            _tempFolder = new VbaTemporaryFolder(so);
            _repositoryFolder = new VbaRepositoryFolder(so, session, sessionSettings);
        }

        public ILocateModules RepositoryFolderModuleLocator => _repositoryFolder.GetModuleLocator();
        public ILocateModules TemporaryFolderModuleLocator => _tempFolder.GetModuleLocator();

        public void Apply(IEnumerable<Patch> patches)
        {
            if (_session.Action == ActionType.Extract)
            {
                ApplyPatches(patches, _tempFolder, _repositoryFolder);
            }
            else
            {
                ApplyPatches(patches, _repositoryFolder, _tempFolder);
                _sessionSettings.BeforePublishHook?.Execute(_tempFolder.FolderPath);
                _tempFolder.Write(_session.FilePath);
            }
        }

        public void Dispose() => _tempFolder.Dispose();

        public IEnumerable<Patch> GetPatches()
        {
            var folderModules = _repositoryFolder.GetCodeModules();
            _tempFolder.Read(_session.FilePath);
            _sessionSettings.AfterExtractHook?.Execute(_tempFolder.FolderPath);
            _tempFolder.FixCase(folderModules.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Item1));
            var vbaFileModules = _tempFolder.GetCodeModules();
            foreach (var patch in Lib.GetModulePatches(_so, _session, _sessionSettings,
                _repositoryFolder.GetModuleLocator(), folderModules,
                _tempFolder.GetModuleLocator(), vbaFileModules))
            {
                yield return patch;
            }
            var projPatch = Lib.GetProjectPatch(_so, _session, _tempFolder.FolderPath);
            if (projPatch != null)
            {
                yield return projPatch;
            }
            var licensesPatch = Lib.GetLicensesPatch(_so, _session, _tempFolder.FolderPath);
            if (licensesPatch != null)
            {
                yield return licensesPatch;
            }
        }

        private void ApplyPatches(IEnumerable<Patch> patches, VbaFolder source, VbaFolder destination)
        {
            foreach (var patch in patches)
            {
                switch (patch.ChangeType)
                {
                    case ChangeType.AddFile:
                        destination.AddModule(patch.ModuleName, patch.ModuleType, source);
                        break;
                    case ChangeType.DeleteFile:
                        destination.DeleteModule(patch.ModuleName, patch.ModuleType);
                        break;
                    case ChangeType.ChangeFormControls:
                        destination.ReplaceFormControls(patch.ModuleName, source);
                        break;
                    default:
                        destination.ReplaceTextModule(patch.ModuleName, patch.ModuleType, source, patch.SideBySideNewText);
                        break;
                }
            }
        }
    }
}
