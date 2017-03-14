using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Text;

namespace VBASync.Model
{
    public class IniFile
    {
        protected readonly Dictionary<string, string> Dict = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);

        private readonly Encoding _encoding;

        public IniFile(IsolatedStorageFile store, string fileName, Encoding encoding = null)
        {
            _encoding = encoding ?? Encoding.Default;
            AddFile(store, fileName);
        }

        public IniFile(string filePath, Encoding encoding = null)
        {
            _encoding = encoding ?? Encoding.Default;
            AddFile(filePath);
        }

        public void AddFile(IsolatedStorageFile store, string fileName)
        {
            if (!store.FileExists(fileName))
            {
                return;
            }
            string s;
            using (var strm = new IsolatedStorageFileStream(fileName, FileMode.Open, FileAccess.Read, store))
            {
                var buf = new byte[strm.Length];
                strm.Read(buf, 0, (int)strm.Length);
                s = _encoding.GetString(buf);
            }
            ProcessString(s);
        }

        public void AddFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }
            ProcessString(File.ReadAllText(filePath, _encoding));
        }

        public void Delete(string subject)
        {
            foreach (var s in Dict.Keys.Where(s => s.StartsWith($"{subject}\0", StringComparison.InvariantCultureIgnoreCase)).ToArray())
            {
                Dict.Remove(s);
            }
        }

        public void Delete(string subject, string key)
        {
            Dict.Remove($"{subject}\0{key}");
        }

        public bool? GetBool(string subject, string key)
        {
            const StringComparison iic = StringComparison.InvariantCultureIgnoreCase;
            if (Dict.TryGetValue($"{subject}\0{key}", out var s))
            {
                if (s.Equals("TRUE", iic) || s.Equals("YES", iic) || s == "1")
                {
                    return true;
                }
                if (s.Equals("FALSE", iic) || s.Equals("NO", iic) || s == "0")
                {
                    return false;
                }
            }
            return null;
        }

        public Guid? GetGuid(string subject, string key)
        {
            if (Dict.TryGetValue($"{subject}\0{key}", out var s) && Guid.TryParse(s, out var u))
            {
                return u;
            }
            return null;
        }

        public int? GetInt(string subject, string key)
        {
            if (Dict.TryGetValue($"{subject}\0{key}", out var s) && (int.TryParse(s, out var i)
                || ((s?.Length ?? 0) > 2 && int.TryParse(s?.Substring(1, s.Length - 2), out i))))
            {
                return i;
            }
            return null;
        }

        public string GetString(string subject, string key)
        {
            if (Dict.TryGetValue($"{subject}\0{key}", out string s) && s != null && s.Length >= 2
                && s[0] == '"' && s[s.Length - 1] == '"')
            {
                return s.Substring(1, s.Length - 2);
            }
            return s;
        }

        public Version GetVersion(string subject, string key)
        {
            if (Dict.TryGetValue($"{subject}\0{key}", out string s) && Version.TryParse(s, out Version v))
            {
                return v;
            }
            return null;
        }

        private void ProcessString(string s)
        {
            var sr = new StringReader(s);
            var subject = "General";
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                line = line.Trim();
                if (line.StartsWith("["))
                {
                    subject = line.Substring(1, line.Length - 2);
                }
                else if (line.Length > 0)
                {
                    var idx = line.IndexOf('=');
                    if (idx == line.Length - 1)
                    {
                        Dict.Remove($"{subject}\0{line.Substring(0, idx).TrimEnd()}");
                    }
                    else if (idx > 0)
                    {
                        Dict[$"{subject}\0{line.Substring(0, idx).TrimEnd()}"] = line.Substring(idx + 1).TrimStart();
                    }
                    else
                    {
                        Dict[$"{subject}\0{line}"] = "";
                    }
                }
            }
        }
    }
}
