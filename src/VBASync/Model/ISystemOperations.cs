using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace VBASync.Model
{
    internal interface ISystemOperations
    {
        bool IsWindows { get; }

        Stream CreateNewFile(string path);
        void DirectoryCreateDirectory(string path);
        void DirectoryDelete(string path, bool recursive);
        IEnumerable<string> DirectoryGetFiles(string folderPath);
        IEnumerable<string> DirectoryGetFiles(string folderPath, string mask);
        IEnumerable<string> DirectoryGetFiles(string folderPath, string mask, bool recurse);
        void FileCopy(string src, string dest);
        void FileCopy(string src, string dest, bool overwrite);
        void FileDelete(string path);
        bool FileExists(string path);
        byte[] FileReadAllBytes(string path);
        string FileReadAllText(string path, Encoding encoding);
        void FileWriteAllBytes(string path, byte[] bytes);
        void FileWriteAllLines(string path, IEnumerable<string> lines, Encoding encoding);
        void FileWriteAllText(string path, string text, Encoding encoding);
        Stream OpenFileForRead(string path);
        Stream OpenFileForWrite(string path);
        string PathCombine(params string[] parts);
        string PathGetDirectoryName(string path);
        string PathGetExtension(string path);
        string PathGetFileName(string path);
        string PathGetFileNameWithoutExtension(string path);
        string PathGetTempPath();
        void ProcessStartAndWaitForExit(ProcessStartInfo psi);
    }
}
