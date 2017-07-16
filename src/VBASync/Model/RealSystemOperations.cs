using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace VBASync.Model
{
    internal class RealSystemOperations : ISystemOperations
    {
        public bool IsWindows => Environment.OSVersion.Platform == PlatformID.Win32NT;

        public Stream CreateNewFile(string path) => File.Create(path);
        public void DirectoryCreateDirectory(string path) => Directory.CreateDirectory(path);
        public void DirectoryDelete(string path, bool recursive) => Directory.Delete(path, recursive);
        public IEnumerable<string> DirectoryGetFiles(string folderPath) => Directory.GetFiles(folderPath);
        public IEnumerable<string> DirectoryGetFiles(string folderPath, string mask) => Directory.GetFiles(folderPath, mask);

        public IEnumerable<string> DirectoryGetFiles(string folderPath, string mask, bool recurse)
            => Directory.GetFiles(folderPath, mask, recurse ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);

        public void FileCopy(string src, string dest) => File.Copy(src, dest);
        public void FileCopy(string src, string dest, bool overwrite) => File.Copy(src, dest, overwrite);
        public void FileDelete(string path) => File.Delete(path);
        public bool FileExists(string path) => File.Exists(path);
        public byte[] FileReadAllBytes(string path) => File.ReadAllBytes(path);
        public string FileReadAllText(string path, Encoding encoding) => File.ReadAllText(path, encoding);
        public void FileWriteAllBytes(string path, byte[] bytes) => File.WriteAllBytes(path, bytes);

        public void FileWriteAllLines(string path, IEnumerable<string> lines, Encoding encoding)
            => File.WriteAllLines(path, lines, encoding);

        public void FileWriteAllText(string path, string text, Encoding encoding) => File.WriteAllText(path, text, encoding);
        public Stream OpenFileForRead(string path) => new FileStream(path, FileMode.Open, FileAccess.Read);
        public Stream OpenFileForWrite(string path) => new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        public string PathCombine(params string[] parts) => Path.Combine(parts);
        public string PathGetDirectoryName(string path) => Path.GetDirectoryName(path);
        public string PathGetExtension(string path) => Path.GetExtension(path);
        public string PathGetFileName(string path) => Path.GetFileName(path);
        public string PathGetFileNameWithoutExtension(string path) => Path.GetFileNameWithoutExtension(path);
        public string PathGetTempPath() => Path.GetTempPath();
        public void ProcessStartAndWaitForExit(ProcessStartInfo psi) => Process.Start(psi).WaitForExit();
    }
}
