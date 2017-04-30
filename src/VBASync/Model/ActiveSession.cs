using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VBASync.Model
{
    public sealed class ActiveSession : IDisposable
    {
        private readonly ISession _session;
        private readonly VbaFolder _vf;

        public ActiveSession(ISession session)
        {
            _session = session;
            _vf = new VbaFolder();
        }

        public string FolderPath => _vf.FolderPath;

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
                            File.Delete(Path.Combine(_session.FolderPath, fileName));
                            if (p.ModuleType == ModuleType.Form)
                            {
                                File.Delete(Path.Combine(_session.FolderPath, p.ModuleName + ".frx"));
                            }
                            break;
                        case ChangeType.ChangeFormControls:
                            File.Copy(Path.Combine(_vf.FolderPath, p.ModuleName + ".frx"), Path.Combine(_session.FolderPath, p.ModuleName + ".frx"), true);
                            break;
                        case ChangeType.Licenses:
                            if (!File.Exists(Path.Combine(_vf.FolderPath, fileName)))
                            {
                                File.Delete(Path.Combine(_session.FolderPath, fileName));
                            }
                            else
                            {
                                File.Copy(Path.Combine(_vf.FolderPath, fileName), Path.Combine(_session.FolderPath, fileName), true);
                            }
                            break;
                        default:
                            File.Copy(Path.Combine(_vf.FolderPath, fileName), Path.Combine(_session.FolderPath, fileName), true);
                            if (p.ChangeType == ChangeType.AddFile && p.ModuleType == ModuleType.Form)
                            {
                                File.Copy(Path.Combine(_vf.FolderPath, p.ModuleName + ".frx"), Path.Combine(_session.FolderPath, p.ModuleName + ".frx"), true);
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
                            File.Delete(Path.Combine(_vf.FolderPath, fileName));
                            if (p.ModuleType == ModuleType.Form)
                            {
                                File.Delete(Path.Combine(_vf.FolderPath, p.ModuleName + ".frx"));
                            }
                            break;
                        case ChangeType.ChangeFormControls:
                            File.Copy(Path.Combine(_session.FolderPath, p.ModuleName + ".frx"), Path.Combine(_vf.FolderPath, p.ModuleName + ".frx"), true);
                            break;
                        case ChangeType.Licenses:
                            if (!File.Exists(Path.Combine(_session.FolderPath, fileName)))
                            {
                                File.Delete(Path.Combine(_vf.FolderPath, fileName));
                            }
                            else
                            {
                                File.Copy(Path.Combine(_session.FolderPath, fileName), Path.Combine(_vf.FolderPath, fileName), true);
                            }
                            break;
                        default:
                            File.Copy(Path.Combine(_session.FolderPath, fileName), Path.Combine(_vf.FolderPath, fileName), true);
                            if (p.ChangeType == ChangeType.AddFile && p.ModuleType == ModuleType.Form)
                            {
                                File.Copy(Path.Combine(_session.FolderPath, p.ModuleName + ".frx"), Path.Combine(_vf.FolderPath, p.ModuleName + ".frx"), true);
                            }
                            break;
                    }
                }
                _vf.Write(_session.FilePath);
            }
        }

        public void Dispose() => _vf.Dispose();

        public IEnumerable<Patch> GetPatches()
        {
            var folderModules = Lib.GetFolderModules(_session.FolderPath);
            _vf.Read(_session.FilePath, folderModules.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Item1));
            foreach (var patch in Lib.GetModulePatches(_session, _vf.FolderPath, folderModules, _vf.ModuleTexts.ToList()))
            {
                yield return patch;
            }
            var projPatch = Lib.GetProjectPatch(_session, _vf.FolderPath);
            if (projPatch != null)
            {
                yield return projPatch;
            }
            var licensesPatch = Lib.GetLicensesPatch(_session, _vf.FolderPath);
            if (licensesPatch != null)
            {
                yield return licensesPatch;
            }
        }
    }
}
