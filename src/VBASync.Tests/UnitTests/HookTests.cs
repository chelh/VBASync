using NUnit.Framework;
using VBASync.Model;

namespace VBASync.Tests.UnitTests
{
    [TestFixture]
    public class HookTests
    {
        [Test]
        public void HookLinux()
        {
            var hook = new Hook("./hook.sh \"{TargetDir}\"", false);
            var psi = hook.GetProcessStartInfo("/tmp/ExtractedVbaFolder/");
            Assert.That(psi.FileName, Is.EqualTo("sh"));
            Assert.That(psi.Arguments, Is.EqualTo("-c \"./hook.sh \\\"/tmp/ExtractedVbaFolder/\\\"\""));
        }

        [Test]
        public void HookWindows()
        {
            var hook = new Hook("hook.bat \"{TargetDir}\"", true);
            var psi = hook.GetProcessStartInfo("C:\\Temp\\ExtractedVbaFolder\\");
            Assert.That(psi.FileName, Is.EqualTo("cmd.exe").IgnoreCase);
            Assert.That(psi.Arguments, Is.EqualTo("/c hook.bat \"C:\\Temp\\ExtractedVbaFolder\\\""));
        }
    }
}
