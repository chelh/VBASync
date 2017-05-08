using System;
using System.Diagnostics;

namespace VBASync.Model
{
    public class Hook
    {
        public Hook(string content)
        {
            Content = content;
        }

        public string Content { get; }

        public void Execute(string targetDir)
        {
            if (string.IsNullOrEmpty(Content))
            {
                return;
            }

            var proc = IsWindows() ? ExecWindows(targetDir) : ExecUnix(targetDir);
            proc.WaitForExit();
        }

        private Process ExecWindows(string targetDir)
        {
            var cmd = Content.Replace("{Target}", targetDir);
            return Process.Start("cmd.exe", $"/c {cmd}");
        }

        private Process ExecUnix(string targetDir)
        {
            var cmd = Content.Replace("{Target}", targetDir);
            return Process.Start("sh", $"-c \"{cmd.Replace("\"", "\\\"")}\"");
        }

        private bool IsWindows() => Environment.OSVersion.Platform == PlatformID.Win32NT;
    }
}
