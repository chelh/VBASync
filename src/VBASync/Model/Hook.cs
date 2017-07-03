using System.Diagnostics;

namespace VBASync.Model
{
    public class Hook
    {
        private readonly ISystemOperations _so;

        public Hook(string content) : this(new RealSystemOperations(), content)
        {
        }

        internal Hook(ISystemOperations so, string content)
        {
            _so = so;
            Content = content;
        }

        public string Content { get; }

        public void Execute(string targetDir)
        {
            if (string.IsNullOrEmpty(Content))
            {
                return;
            }

            _so.ProcessStartAndWaitForExit(GetProcessStartInfo(targetDir));
        }

        internal ProcessStartInfo GetProcessStartInfo(string targetDir)
        {
            if (string.IsNullOrEmpty(Content))
            {
                return null;
            }

            if (_so.IsWindows)
            {
                return new ProcessStartInfo("cmd.exe", $"/c {Content.Replace("{TargetDir}", targetDir)}");
            }
            else
            {
                return new ProcessStartInfo("sh", $"-c \"{Content.Replace("\"", "\\\"").Replace("{TargetDir}", targetDir)}\"");
            }
        }
    }
}
