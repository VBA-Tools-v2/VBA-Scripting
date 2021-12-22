Attribute VB_Name = "test_FileSystem"
''
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@testmodule
'@folder Tests
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Private Module

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type TTest
    Assert As Object
    Fakes As Object
    ScrFileSystem As Object             ' Scripting.FileSystemObject
    VbaFileSystem As FileSystemObject   ' VBA.FileSystemObject
End Type

Private This As TTest

' ============================================= '
' Test Methods
' ============================================= '

' --------------------------------------------- '
' BuildPath
' --------------------------------------------- '

'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathEmpty_NameEmpty()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath(vbNullString, vbNullString), This.VbaFileSystem.BuildPath(vbNullString, vbNullString)
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathEmpty_NameValid()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath(vbNullString, "Hello World"), This.VbaFileSystem.BuildPath(vbNullString, "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathValid_NameEmpty()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Hello World", vbNullString), This.VbaFileSystem.BuildPath("Hello World", vbNullString)
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathRelatvie_NameValid()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("..", "Hello World"), This.VbaFileSystem.BuildPath("..", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathEndsDoubleSeparator_NameValid()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("C:\Users\JohnDoe\Documents\\", "Hello World"), This.VbaFileSystem.BuildPath("C:\Users\JohnDoe\Documents\\", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameNoSeparator()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", "Hello World"), This.VbaFileSystem.BuildPath("Documents", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameSeparatorBackslash()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", "\Hello World"), This.VbaFileSystem.BuildPath("Documents", "\Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameSeparatorForwardslash()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", "/Hello World"), This.VbaFileSystem.BuildPath("Documents", "/Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameSeparatorColon()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", ":Hello World"), This.VbaFileSystem.BuildPath("Documents", ":Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorBackslash_NameNoSeparator()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("C:\Users\JohnDoe\Documents", "Hello World"), This.VbaFileSystem.BuildPath("C:\Users\JohnDoe\Documents", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorForwardslash_NameNoSeparator()
    This.Assert.AreEqual "https://www.google.co.nz/HelloWorld", This.VbaFileSystem.BuildPath("https://www.google.co.nz", "HelloWorld")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorColon_NameNoSeparator()
    This.Assert.AreEqual This.ScrFileSystem.BuildPath(":Documents", "Hello World"), This.VbaFileSystem.BuildPath(":Documents", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorForwardslash_NameSeparatorBackslash()
    This.Assert.AreEqual "https://www.google.co.nz/HelloWorld", This.VbaFileSystem.BuildPath("https://www.google.co.nz", "\HelloWorld")
End Sub
End Sub

' --------------------------------------------- '
' GetBaseName
' --------------------------------------------- '

'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_EmptyString()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName(vbNullString), This.VbaFileSystem.GetBaseName(vbNullString), "GetBaseName does not produce expected result when passed an empty string."
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_InvalidString()
'---ARRANGE---
    Dim test_InvalidPath As String
'---ACT---
    test_InvalidPath = "LoremIpsum"
'---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName(test_InvalidPath), This.VbaFileSystem.GetBaseName(test_InvalidPath), "GetBaseName does not produce expected result when passed an invalid file path."
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_RelativePath()
'---ARRANGE---
    Dim test_RelativePath As String
'---ACT---
    test_RelativePath = "..\Documents\HelloWorld.txt"
'---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName(test_RelativePath), This.VbaFileSystem.GetBaseName(test_RelativePath), "GetBaseName does not produce expected result when passed a relative file path."
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_AbsolutePath()
'---ARRANGE---
    Dim test_AbsolutePath As String
'---ACT---
    test_AbsolutePath = "C:\Users\JohnDoe\Documents\HelloWorld.txt"
'---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName(test_AbsolutePath), This.VbaFileSystem.GetBaseName(test_AbsolutePath), "GetBaseName does not produce expected result when passed an absolute file path."
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_PathSeparatorForwardslash()

End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_PathSeparatorColon()

End Sub
' --------------------------------------------- '
' Speed Tests
' --------------------------------------------- '

'@testmethod FileSystem.SpeedTest
Private Sub speedtest_BuildPath()
    Dim test_Temp As String
    Dim test_Long As Long
    Dim test_StartTime As Date
    Dim test_FinishTime As Date
    Dim test_VbaMS As Double
    Dim test_ScrMs As Double
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 1000000
        test_Temp = This.ScrFileSystem.BuildPath("C:\Users\JohnDoe\Documents", "Hello World")
    Next test_Long
    test_FinishTime = VBA.Date + CDate(VBA.Timer / 86400)
    test_ScrMs = VBA.Round((test_FinishTime - test_StartTime) * 86400 * 1000, 4)
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 1000000
        test_Temp = This.VbaFileSystem.BuildPath("C:\Users\JohnDoe\Documents", "Hello World")
    Next test_Long
    test_FinishTime = VBA.Date + CDate(VBA.Timer / 86400)
    test_VbaMS = VBA.Round((test_FinishTime - test_StartTime) * 86400 * 1000, 4)
    
    This.Assert.Inconclusive "SCR=" & test_ScrMs & "ms | VBA=" & test_VbaMS & "ms | " & VBA.IIf(test_VbaMS > test_ScrMs, "Scripting", "VBA") & " is " & VBA.Round(VBA.IIf(test_VbaMS > test_ScrMs, test_VbaMS / test_ScrMs, test_ScrMs / test_VbaMS), 4) & " times faster."
End Sub
' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

'@TestInitialize
Private Sub TestInitialize()
    ' This method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' This method runs after every test in the module.
End Sub

'@ModuleInitialize
Private Sub ModuleInitialize()
    With This
        Set .Assert = CreateObject("Rubberduck.AssertClass")
        Set .Fakes = CreateObject("Rubberduck.FakesProvider")
        Set .ScrFileSystem = CreateObject("Scripting.FileSystemObject")
        Set .VbaFileSystem = New FileSystemObject
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    With This
        Set .Assert = Nothing
        Set .Fakes = Nothing
        Set .ScrFileSystem = Nothing
        Set .VbaFileSystem = Nothing
    End With
End Sub
