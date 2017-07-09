using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VBASync.Model
{
    public class VbaFolder
    {
        private readonly ISystemOperations _so;

        public VbaFolder(string folderPath) : this(new RealSystemOperations(), folderPath)
        {
        }

        internal VbaFolder(ISystemOperations so, string folderPath)
        {
            _so = so;
            FolderPath = folderPath;
        }

        public string FolderPath { get; }
        public Encoding ProjectEncoding { get; protected set; }

        public IList<KeyValuePair<string, Tuple<string, ModuleType>>> GetModules()
        {
            var modulesText = new Dictionary<string, Tuple<string, ModuleType>>();
            var extensions = new[] { ".bas", ".cls", ".frm" };
            var projIni = new ProjectIni(_so.PathCombine(FolderPath, "Project.ini"));
            projIni.AddFile(_so.PathCombine(FolderPath, "Project.ini.local"));
            ProjectEncoding = Encoding.GetEncoding(projIni.GetInt("General", "CodePage") ?? Encoding.Default.CodePage);
            // no need to re-read projIni, since we only wanted to get the encoding off of it anyway!
            foreach (var filePath in GetModuleFilePaths())
            {
                var moduleText = _so.FileReadAllText(filePath, ProjectEncoding).TrimEnd('\r', '\n') + "\r\n";
                modulesText[_so.PathGetFileNameWithoutExtension(filePath)] = Tuple.Create(moduleText, ModuleProcessing.TypeFromText(moduleText));
            }
            return modulesText.ToList();
        }

        protected virtual IEnumerable<string> GetModuleFilePaths() => GetModuleFilePaths(false);

        protected IEnumerable<string> GetModuleFilePaths(bool recurse)
        {
            return _so.DirectoryGetFiles(FolderPath, "*.bas", recurse)
                .Concat(_so.DirectoryGetFiles(FolderPath, "*.cls", recurse))
                .Concat(_so.DirectoryGetFiles(FolderPath, "*.frm", recurse))
                .Select(s => _so.PathCombine(FolderPath, _so.PathGetFileName(s)))
                .OrderBy(s => s);
        }
    }
}
