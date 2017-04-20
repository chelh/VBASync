using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBASync.Model;

namespace VBASync.Tests.UnitTests
{
    [TestClass]
    public class FixCaseTests
    {
        [TestMethod]
        public void FixCaseBasic()
        {
            const string oldFile = @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ThisWorkbook""
Attribute VB_Base = ""0{00020819-0000-0000-C000-000000000046}""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' hello world!
Sub tes()
    Dim HelloWorld$
    HelloWorld = ""Hello, world!""
    MsgBox HelloWorld
End Sub
Sub tes2()
    MsgBox 2 + 2
End Sub
Sub tes3()
    MsgBox ""Hello, world!""
End Sub
";
            const string newFileRaw = @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ThisWorkbook""
Attribute VB_Base = ""0{00020819-0000-0000-C000-000000000046}""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Hello world!
Sub Tes()
    Dim helloworld$
    helloworld = ""Hello, World!""
    MsgBox helloworld
End Sub

Sub Tes2()
    MsgBox 2 + 3
End Sub

Sub Tes3()
    MsgBox ""Hello, world!""
End Sub
";
            const string newFileFix = @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ThisWorkbook""
Attribute VB_Base = ""0{00020819-0000-0000-C000-000000000046}""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Hello world!
Sub tes()
    Dim HelloWorld$
    helloworld = ""Hello, World!""
    MsgBox HelloWorld
End Sub

Sub tes2()
    MsgBox 2 + 3
End Sub

Sub tes3()
    MsgBox ""Hello, world!""
End Sub
";
            Assert.AreEqual(newFileFix, ModuleProcessing.FixCase(oldFile, newFileRaw));
        }
    }
}
