using System;
using System.Diagnostics;

namespace VBASync.Model
{
    public class Hook
    {
        private readonly bool _isWindows;

        public Hook(string content) : this(content, Environment.OSVersion.Platform == PlatformID.Win32NT)
        {
        }

        internal Hook(string content, bool isWindows)
        {
            Content = content;
            _isWindows = isWindows;
        }

        public string Content { get; }

        public void Execute(string targetDir)
        {
            if (string.IsNullOrEmpty(Content))
            {
                return;
            }

            Process.Start(GetProcessStartInfo(targetDir)).WaitForExit();
        }

        internal ProcessStartInfo GetProcessStartInfo(string targetDir)
        {
            if (string.IsNullOrEmpty(Content))
            {
                return null;
            }

            if (_isWindows)
            {
                return new ProcessStartInfo("cmd.exe", $"/c {Content}")
                {
                    WorkingDirectory = targetDir
                };
            }
            else
            {
                return new ProcessStartInfo("sh", $"-c \"{Content.Replace("\"", "\\\"")}\"")
                {
                    WorkingDirectory = targetDir
                };
            }
        }
    }
}
