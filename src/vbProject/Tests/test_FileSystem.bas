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
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath(vbNullString, vbNullString), This.VbaFileSystem.BuildPath(vbNullString, vbNullString)
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathEmpty_NameValid()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath(vbNullString, "Hello World"), This.VbaFileSystem.BuildPath(vbNullString, "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathValid_NameEmpty()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Hello World", vbNullString), This.VbaFileSystem.BuildPath("Hello World", vbNullString)
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathEndsDoubleSeparator_NameValid()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("C:\Users\JohnDoe\Documents\\", "Hello World"), This.VbaFileSystem.BuildPath("C:\Users\JohnDoe\Documents\\", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameNoSeparator()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", "Hello World"), This.VbaFileSystem.BuildPath("Documents", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameSeparatorBackSlash()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", "\Hello World"), This.VbaFileSystem.BuildPath("Documents", "\Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameSeparatorForwardSlash()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", "/Hello World"), This.VbaFileSystem.BuildPath("Documents", "/Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathNoSeparator_NameSeparatorColon()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("Documents", ":Hello World"), This.VbaFileSystem.BuildPath("Documents", ":Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorBackSlash_NameNoSeparator()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("\Documents", "Hello World"), This.VbaFileSystem.BuildPath("\Documents", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorForwardSlash_NameNoSeparator()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("https://www.google.co.nz", "Hello World"), This.VbaFileSystem.BuildPath("https://www.google.co.nz", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorColon_NameNoSeparator()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath(":Documents", "Hello World"), This.VbaFileSystem.BuildPath(":Documents", "Hello World")
End Sub
'@testmethod FileSystem.BuildPath
Private Sub BuildPath_PathSeparatorForwardSlash_NameSeparatorBackwardsSlash()
    '---ASSERT---
    This.Assert.AreEqual This.ScrFileSystem.BuildPath("https://www.google.co.nz", "\Hello World"), This.VbaFileSystem.BuildPath("https://www.google.co.nz", "\Hello World")
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
