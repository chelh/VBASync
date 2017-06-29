using NUnit.Framework;
using System.Text;
using VBASync.Model;

namespace VBASync.Tests.UnitTests
{
    [TestFixture]
    public class ModuleProcessingTests
    {
        [Test]
        public void StubOutBasic()
        {
            var sb = new StringBuilder();
            sb.Append("VERSION 1.0 CLASS\r\n");
            sb.Append("BEGIN\r\n");
            sb.Append("  MultiUse = -1  'True\r\n");
            sb.Append("END\r\n");
            sb.Append("Attribute VB_Name = \"ThisWorkbook\"\r\n");
            sb.Append("Attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"\r\n");
            sb.Append("Attribute VB_GlobalNameSpace = False\r\n");
            sb.Append("Attribute VB_Creatable = False\r\n");
            sb.Append("Attribute VB_PredeclaredId = True\r\n");
            sb.Append("Attribute VB_Exposed = True\r\n");
            sb.Append("Attribute VB_TemplateDerived = False\r\n");
            sb.Append("Attribute VB_Customizable = True\r\n");
            var expected = sb.ToString();
            sb.Append("Option Explicit\r\n\r\n");
            sb.Append("Private Sub Workbook_Open()\r\n");
            sb.Append("    MsgBox \"Hello World!\"\r\n");
            sb.Append("End Sub\r\n");
            Assert.That(ModuleProcessing.StubOut(sb.ToString()), Is.EqualTo(expected));
        }
    }
}
