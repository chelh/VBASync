using NUnit.Framework;
using System;
using VBASync.Model;

namespace VBASync.Tests.UnitTests
{
    [TestFixture]
    public class IniFileTests
    {
        [Test]
        public void IniDefaultBoolIsNull()
        {
            var content = "[General]" + Environment.NewLine;
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.Null);
        }

        [Test]
        public void IniDefaultSubject()
        {
            var content = "BoolTest=0" + Environment.NewLine;
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.EqualTo(false));
        }

        [Test]
        public void IniInvalidBoolIsNull()
        {
            var content = "[General]" + Environment.NewLine
                + "BoolTest=bubba" + Environment.NewLine;
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.Null);
        }

        [Test]
        public void IniOverride()
        {
            var content = "[General]" + Environment.NewLine
                + "BoolTest=0" + Environment.NewLine
                + "BoolTest=1" + Environment.NewLine;
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.EqualTo(true));
        }

        [TestCase("0", ExpectedResult = false)]
        [TestCase("1", ExpectedResult = true)]
        [TestCase("False", ExpectedResult = false)]
        [TestCase("FALSE", ExpectedResult = false)]
        [TestCase("No", ExpectedResult = false)]
        [TestCase("NO", ExpectedResult = false)]
        [TestCase("True", ExpectedResult = true)]
        [TestCase("TRUE", ExpectedResult = true)]
        [TestCase("Yes", ExpectedResult = true)]
        [TestCase("YES", ExpectedResult = true)]
        public bool? IniParsesBool(string value)
        {
            var content = "[General]" + Environment.NewLine
                + "BoolTest=" + value + Environment.NewLine;
            return MakeIni(content).GetBool("General", "BoolTest");
        }

        [Test]
        public void IniParsesGuid()
        {
            var content = "[General]" + Environment.NewLine
                + "GuidTest=b5b92f29-a4e3-4da0-8a93-608143212733" + Environment.NewLine;
            Assert.That(MakeIni(content).GetGuid("General", "GuidTest"),
                Is.EqualTo(new Guid("b5b92f29-a4e3-4da0-8a93-608143212733")));
        }

        [TestCase(@"""-7""", ExpectedResult = -7)]
        [TestCase(@"""5""", ExpectedResult = 5)]
        [TestCase("-23", ExpectedResult = -23)]
        [TestCase("4", ExpectedResult = 4)]
        public int? IniParsesInt(string value)
        {
            var content = "[General]" + Environment.NewLine
                + "IntTest=" + value + Environment.NewLine;
            return MakeIni(content).GetInt("General", "IntTest");
        }

        [TestCase(@" ""value with spaces inside  ""   ", ExpectedResult = "value with spaces inside  ")]
        [TestCase("bubba", ExpectedResult = "bubba")]
        public string IniParsesString(string value)
        {
            var content = "[General]" + Environment.NewLine
                + "StringTest=" + value + Environment.NewLine;
            return MakeIni(content).GetString("General", "StringTest");
        }

        [Test]
        public void IniParsesVersion()
        {
            var content = "[General]" + Environment.NewLine
                + "VersionTest=215397.2" + Environment.NewLine;
            Assert.That(MakeIni(content).GetVersion("General", "VersionTest"),
                Is.EqualTo(new Version(215397, 2)));
        }

        [Test]
        public void IniParsesSpacedBool()
        {
            var content = "  [General] " + Environment.NewLine
                + "    BoolTest = 1  " + Environment.NewLine;
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.EqualTo(true));
        }

        [Test]
        public void IniSubjectIsCaseInsensitive()
        {
            var content = "[general]" + Environment.NewLine
                + "BoolTest=1" + Environment.NewLine;
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.EqualTo(true));
        }

        [Test]
        public void IniValueIsCaseInsensitive()
        {
            var content = "[General]" + Environment.NewLine
                + "booltest=1" + Environment.NewLine;
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.EqualTo(true));
        }

        [Test]
        public void IniWithNoNewLine()
        {
            Assert.That(MakeIni("BoolTest=1").GetBool("General", "BoolTest"), Is.EqualTo(true));
        }

        [Test]
        public void IniWithoutEndingNewLine()
        {
            var content = "[General]" + Environment.NewLine
                + "BoolTest=1";
            Assert.That(MakeIni(content).GetBool("General", "BoolTest"), Is.EqualTo(true));
        }

        private IniFile MakeIni(string content)
        {
            var ini = new IniFile();
            ini.ProcessString(content);
            return ini;
        }
    }
}
