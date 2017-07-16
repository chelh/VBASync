using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using VBASync.Model;

namespace VBASync.Tests.Mocks
{
    internal class FakeSystemOperations : ISystemOperations
    {
        private readonly ConcurrentDictionary<string, byte[]> _files
            = new ConcurrentDictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        public virtual bool IsWindows => false;

        public virtual Stream CreateNewFile(string path)
        {
            if (!_files.TryAdd(path, new byte[0]))
            {
                throw new ApplicationException($"Tried to create new file at '{path}', which already exists!");
            }
            return new MemoryStreamWithFlushCallback(s => _files[path] = s.ToArray());
        }

        public virtual void DirectoryCreateDirectory(string path)
        {
        }

        public virtual void DirectoryDelete(string path, bool recursive)
        {
        }

        public virtual IEnumerable<string> DirectoryGetFiles(string folderPath)
            => DirectoryGetFiles(folderPath, "*", false);

        public virtual IEnumerable<string> DirectoryGetFiles(string folderPath, string mask)
            => DirectoryGetFiles(folderPath, mask, false);

        public virtual IEnumerable<string> DirectoryGetFiles(string folderPath, string mask, bool recurse)
        {
            if (!folderPath.EndsWith("/"))
            {
                folderPath += "/";
            }
            foreach (var s in _files.Keys)
            {
                if (InRightDirectory(s) && FitsMask(PathGetFileName(s), mask))
                {
                    yield return s;
                }
            }

            bool FitsMask(string fileName, string fileMask)
            {
                if (fileMask == "*")
                {
                    return true;
                }
                return new Regex("^" + fileMask.Replace(".", "[.]").Replace("*", ".*").Replace("?", ".")
                    + "$", RegexOptions.IgnoreCase).IsMatch(fileName);
            }

            bool InRightDirectory(string filePath)
            {
                if (!filePath.StartsWith(folderPath))
                {
                    return false;
                }
                if (recurse)
                {
                    return true;
                }
                return filePath.LastIndexOf('/') < folderPath.Length;
            }
        }

        public virtual void FileCopy(string src, string dest) => FileCopy(src, dest, false);

        public virtual void FileCopy(string src, string dest, bool overwrite)
        {
            if (overwrite)
            {
                _files[dest] = _files[src];
            }
            else if (!_files.TryAdd(dest, _files[src]))
            {
                throw new ApplicationException($"Tried to copy to a location '{dest}' that already exists!");
            }
        }

        public virtual void FileDelete(string path)
        {
            if (!_files.TryRemove(path, out _))
            {
                throw new ApplicationException($"Tried to remove '{path}', which doesn't exist!");
            }
        }

        public virtual bool FileExists(string path) => _files.ContainsKey(path);
        public virtual byte[] FileReadAllBytes(string path) => _files[path];

        public virtual string FileReadAllText(string path, Encoding encoding)
            => encoding.GetString(FileReadAllBytes(path));

        public virtual void FileWriteAllBytes(string path, byte[] bytes) => _files[path] = bytes;

        public virtual void FileWriteAllLines(string path, IEnumerable<string> lines, Encoding encoding)
            => FileWriteAllText(path, string.Join("\n", lines), encoding);

        public virtual void FileWriteAllText(string path, string text, Encoding encoding)
            => FileWriteAllBytes(path, encoding.GetBytes(text));

        public virtual Stream OpenFileForRead(string path) => new MemoryStream(_files[path]);

        public virtual Stream OpenFileForWrite(string path)
            => new MemoryStreamWithFlushCallback(FileReadAllBytes(path), s => _files[path] = s.ToArray());

        public string PathCombine(params string[] parts)
        {
            var sb = new StringBuilder();
            for (var i = 0; i < parts.Length; ++i)
            {
                if (i > 0)
                {
                    sb.Append('/');
                }
                if (i < parts.Length - 1)
                {
                    sb.Append(parts[i].TrimEnd('/'));
                }
                else
                {
                    sb.Append(parts[i]);
                }
            }
            return sb.ToString();
        }

        public string PathGetDirectoryName(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }
            if (path == "/")
            {
                return "/";
            }
            if (path.EndsWith("/"))
            {
                path = path.Substring(0, path.Length - 1);
            }
            var idx = path.LastIndexOf('/');
            if (idx == 0)
            {
                return "/";
            }
            if (idx < 0)
            {
                return "";
            }
            return path.Substring(0, idx);
        }

        public virtual string PathGetExtension(string path)
        {
            var idx = path.LastIndexOf('.');
            return idx >= 0 ? path.Substring(idx) : "";
        }

        public virtual string PathGetFileName(string path)
        {
            var idx = path.LastIndexOf('/');
            return idx >= 0 ? path.Substring(idx + 1) : path;
        }

        public virtual string PathGetFileNameWithoutExtension(string path)
        {
            var ext = PathGetExtension(path);
            var fn = PathGetFileName(path);
            return fn.Substring(0, fn.Length - ext.Length);
        }

        public virtual string PathGetTempPath() => "/tmp/";

        public virtual void ProcessStartAndWaitForExit(ProcessStartInfo psi)
        {
        }

        private class MemoryStreamWithFlushCallback : MemoryStream
        {
            private readonly Action<MemoryStream> _flushCallback;

            internal MemoryStreamWithFlushCallback(Action<MemoryStream> flushCallback)
            {
                _flushCallback = flushCallback;
            }

            internal MemoryStreamWithFlushCallback(byte[] buffer, Action<MemoryStream> flushCallback) : this(flushCallback)
            {
                Write(buffer, 0, buffer.Length);
                base.Flush();
            }

            public override void Flush()
            {
                base.Flush();
                _flushCallback(this);
            }
        }
    }

    internal class WindowsFakeSystemOperations : FakeSystemOperations
    {
        public override bool IsWindows => true;
    }
}
