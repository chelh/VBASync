using System;
using System.Collections.Generic;
using System.Linq;

namespace VBASync.Model
{
    public sealed class ActiveSession : IDisposable
    {
        private readonly ISystemOperations _so;
        private readonly ISession _session;
        private readonly ISessionSettings _sessionSettings;
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
        }

        public string TemporaryFolderPath => _tempFolder.FolderPath;

        public void Apply(IEnumerable<Patch> changes)
        {
            if (_session.Action == ActionType.Extract)
            {
                foreach (var p in changes)
                {
                    var fileName = p.ModuleName + ModuleProcessing.ExtensionFromType(p.ModuleType);
                    switch (p.ChangeType)
                    {
                        case ChangeType.DeleteFile:
                            _so.FileDelete(_so.PathCombine(_session.FolderPath, fileName));
                            if (p.ModuleType == ModuleType.Form)
                            {
                                _so.FileDelete(_so.PathCombine(_session.FolderPath, p.ModuleName + ".frx"));
                            }
                            break;
                        case ChangeType.ChangeFormControls:
                            _so.FileCopy(_so.PathCombine(_tempFolder.FolderPath, p.ModuleName + ".frx"), _so.PathCombine(_session.FolderPath, p.ModuleName + ".frx"), true);
                            break;
                        case ChangeType.Licenses:
                            if (!_so.FileExists(_so.PathCombine(_tempFolder.FolderPath, fileName)))
                            {
                                _so.FileDelete(_so.PathCombine(_session.FolderPath, fileName));
                            }
                            else
                            {
                                _so.FileCopy(_so.PathCombine(_tempFolder.FolderPath, fileName), _so.PathCombine(_session.FolderPath, fileName), true);
                            }
                            break;
                        default:
                            _so.FileCopy(_so.PathCombine(_tempFolder.FolderPath, fileName), _so.PathCombine(_session.FolderPath, fileName), true);
                            if (p.ChangeType == ChangeType.AddFile && p.ModuleType == ModuleType.Form)
                            {
                                _so.FileCopy(_so.PathCombine(_tempFolder.FolderPath, p.ModuleName + ".frx"), _so.PathCombine(_session.FolderPath, p.ModuleName + ".frx"), true);
                            }
                            break;
                    }
                }
            }
            else
            {
                foreach (var p in changes)
                {
                    var fileName = p.ModuleName + ModuleProcessing.ExtensionFromType(p.ModuleType);
                    switch (p.ChangeType)
                    {
                        case ChangeType.DeleteFile:
                            _so.FileDelete(_so.PathCombine(_tempFolder.FolderPath, fileName));
                            if (p.ModuleType == ModuleType.Form)
                            {
                                _so.FileDelete(_so.PathCombine(_tempFolder.FolderPath, p.ModuleName + ".frx"));
                            }
                            break;
                        case ChangeType.ChangeFormControls:
                            _so.FileCopy(_so.PathCombine(_session.FolderPath, p.ModuleName + ".frx"), _so.PathCombine(_tempFolder.FolderPath, p.ModuleName + ".frx"), true);
                            break;
                        case ChangeType.Licenses:
                            if (!_so.FileExists(_so.PathCombine(_session.FolderPath, fileName)))
                            {
                                _so.FileDelete(_so.PathCombine(_tempFolder.FolderPath, fileName));
                            }
                            else
                            {
                                _so.FileCopy(_so.PathCombine(_session.FolderPath, fileName), _so.PathCombine(_tempFolder.FolderPath, fileName), true);
                            }
                            break;
                        default:
                            if (_so.FileExists(_so.PathCombine(_session.FolderPath, fileName)))
                            {
                                _so.FileCopy(_so.PathCombine(_session.FolderPath, fileName), _so.PathCombine(_tempFolder.FolderPath, fileName), true);
                            }
                            else
                            {
                                _so.FileWriteAllText(_so.PathCombine(_tempFolder.FolderPath, fileName), p.SideBySideNewText, _tempFolder.ProjectEncoding);
                            }
                            if (p.ChangeType == ChangeType.AddFile && p.ModuleType == ModuleType.Form)
                            {
                                _so.FileCopy(_so.PathCombine(_session.FolderPath, p.ModuleName + ".frx"), _so.PathCombine(_tempFolder.FolderPath, p.ModuleName + ".frx"), true);
                            }
                            break;
                    }
                }
                _sessionSettings.BeforePublishHook?.Execute(_tempFolder.FolderPath);
                _tempFolder.Write(_session.FilePath);
            }
        }

        public void Dispose() => _tempFolder.Dispose();

        public IEnumerable<Patch> GetPatches()
        {
            var folderModules = Lib.GetFolderModules(_so, _session.FolderPath);
            _tempFolder.Read(_session.FilePath);
            _sessionSettings.AfterExtractHook?.Execute(_tempFolder.FolderPath);
            _tempFolder.FixCase(folderModules.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Item1));
            var vbaFileModules = Lib.GetFolderModules(_so, _tempFolder.FolderPath);
            foreach (var patch in Lib.GetModulePatches(_so, _session, _sessionSettings, _tempFolder.FolderPath, folderModules, vbaFileModules))
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
    }
}
