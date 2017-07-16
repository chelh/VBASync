using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VBASync.Model
{
    internal abstract class VbaFolder
    {
        protected List<string> ModuleFilePaths;

        private readonly ISystemOperations _so;

        protected internal VbaFolder(ISystemOperations so, string folderPath)
        {
            _so = so;
            FolderPath = folderPath;
        }

        public string FolderPath { get; }
        public Encoding ProjectEncoding { get; protected set; }

        public void AddModule(string name, ModuleType type, VbaFolder source)
        {
            var path = _so.PathCombine(FolderPath, name + ModuleProcessing.ExtensionFromType(type));
            _so.FileCopy(source.FindModulePath(name, type), path);
            if (type == ModuleType.Form)
            {
                var frxPath = path.Substring(0, path.Length - 4) + ".frx";
                _so.FileCopy(source.FindFrxPath(name), frxPath);
            }
            ModuleFilePaths.Add(path);
        }

        public void DeleteModule(string name, ModuleType type)
        {
            var path = FindModulePath(name, type);
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            _so.FileDelete(path);
            if (type == ModuleType.Form)
            {
                var frxPath = FindFrxPath(name);
                if (!string.IsNullOrEmpty(frxPath))
                {
                    _so.FileDelete(frxPath);
                }
            }
            ModuleFilePaths.Remove(path);
        }

        public IList<KeyValuePair<string, Tuple<string, ModuleType>>> GetCodeModules()
        {
            var modulesText = new Dictionary<string, Tuple<string, ModuleType>>();

            var projIni = new ProjectIni(_so.PathCombine(FolderPath, "Project.ini"));
            projIni.AddFile(_so.PathCombine(FolderPath, "Project.ini.local"));
            ProjectEncoding = Encoding.GetEncoding(projIni.GetInt("General", "CodePage") ?? Encoding.Default.CodePage);
            // no need to re-read projIni, since we only wanted to get the encoding off of it anyway!

            ModuleFilePaths = GetModuleFilePaths();
            foreach (var filePath in ModuleFilePaths)
            {
                var ext = _so.PathGetExtension(filePath).ToUpperInvariant();
                if (ext == ".BAS" || ext == ".CLS" || ext == ".FRM")
                {
                    var moduleText = _so.FileReadAllText(filePath, ProjectEncoding).TrimEnd('\r', '\n') + "\r\n";
                    modulesText[_so.PathGetFileNameWithoutExtension(filePath)] = Tuple.Create(moduleText, ModuleProcessing.TypeFromText(moduleText));
                }
            }

            return modulesText.ToList();
        }

        public ILocateModules GetModuleLocator() => new ModuleLocator(this);

        public void ReplaceFormControls(string name, VbaFolder source)
        {
            _so.FileCopy(FindFrxPath(source.FindModulePath(name, ModuleType.Form)),
                FindFrxPath(FindModulePath(name, ModuleType.Form)), true);
        }

        public void ReplaceTextModule(string name, ModuleType type, VbaFolder source, string fallbackText)
        {
            var sourcePath = source.FindModulePath(name, type);
            var destinationPath = FindModulePath(name, type);
            if (string.IsNullOrEmpty(destinationPath))
            {
                destinationPath = _so.PathCombine(FolderPath, name + ModuleProcessing.ExtensionFromType(type));
            }
            if (_so.FileExists(sourcePath))
            {
                if (source.ProjectEncoding.Equals(ProjectEncoding))
                {
                    _so.FileCopy(sourcePath, destinationPath, true);
                }
                else
                {
                    _so.FileWriteAllText(destinationPath, _so.FileReadAllText(sourcePath, source.ProjectEncoding), ProjectEncoding);
                }
            }
            else
            {
                _so.FileWriteAllText(destinationPath, fallbackText, ProjectEncoding);
            }
        }

        protected string FindFrxPath(string name)
        {
            var frmPath = FindModulePath(name, ModuleType.Form);
            if (string.IsNullOrEmpty(frmPath))
            {
                return null;
            }

            var typicalFrxPath = frmPath.Substring(0, frmPath.Length - 4) + ".frx";
            if (_so.FileExists(typicalFrxPath))
            {
                return typicalFrxPath;
            }
            else
            {
                // should only reach here if we're not on Windows, and the .frx filename has peculiar casing
                var siblings = _so.DirectoryGetFiles(_so.PathGetDirectoryName(frmPath), "*.*", false);
                return siblings.FirstOrDefault(s => string.Equals(_so.PathGetExtension(s), ".frx", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(_so.PathGetFileNameWithoutExtension(s), name, StringComparison.OrdinalIgnoreCase));
            }
        }

        protected string FindModulePath(string name, ModuleType type)
            => ModuleFilePaths.Find(s => string.Equals(_so.PathGetFileName(s), name + ModuleProcessing.ExtensionFromType(type),
                StringComparison.OrdinalIgnoreCase));

        protected virtual List<string> GetModuleFilePaths() => GetModuleFilePaths(false);

        protected List<string> GetModuleFilePaths(bool recurse)
        {
            // don't filter on file extension because on Linux DirectoryGetFiles is case-sensitive!
            var allFiles = _so.DirectoryGetFiles(FolderPath, "*.*", recurse);
            var modulesFound = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Func<string, bool> filter = path =>
            {
                var name = _so.PathGetFileNameWithoutExtension(path);
                var ext = _so.PathGetExtension(path).ToUpperInvariant();
                return (ext == ".BAS" || ext == ".CLS" || ext == ".FRM" || ext == ".INI" || ext == ".BIN")
                    && modulesFound.Add(name);
            };
            return allFiles.Where(filter).ToList();
        }

        private class ModuleLocator : ILocateModules
        {
            private readonly VbaFolder _parent;

            internal ModuleLocator(VbaFolder parent)
            {
                _parent = parent;
            }

            public string GetFrxPath(string name) => _parent.FindFrxPath(name);
            public string GetModulePath(string name, ModuleType type) => _parent.FindModulePath(name, type);
        }
    }
}
