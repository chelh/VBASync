using NUnit.Framework;
using System;
using VBASync.Model;

namespace VBASync.Tests.UnitTests
{
    [TestFixture]
    public class AppIniFileTests
    {
        [TestCase("Extract", ExpectedResult = ActionType.Extract)]
        [TestCase("EXTRACT", ExpectedResult = ActionType.Extract)]
        [TestCase("Publish", ExpectedResult = ActionType.Publish)]
        [TestCase("PUBLISH", ExpectedResult = ActionType.Publish)]
        public ActionType? AppIniParsesActionType(string value)
        {
            var content = "[General]" + Environment.NewLine
                + "ActionTypeTest=" + value + Environment.NewLine;
            var ini = new AppIniFile();
            ini.ProcessString(content);
            return ini.GetActionType("General", "ActionTypeTest");
        }
    }
}
