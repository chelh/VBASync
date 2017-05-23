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
            return Process.Start(new ProcessStartInfo("cmd.exe", $"/c {Content}")
            {
                WorkingDirectory = targetDir
            });
        }

        private Process ExecUnix(string targetDir)
        {
            return Process.Start(new ProcessStartInfo("sh", $"-c \"{Content.Replace("\"", "\\\"")}\"")
            {
                WorkingDirectory = targetDir
            });
        }

        private bool IsWindows() => Environment.OSVersion.Platform == PlatformID.Win32NT;
    }
}
