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
    TestFolderPath As String            ' Path to a folder created for these tests.
    TestFilePath As String              ' Path to a .txt file created for these tests.
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

' --------------------------------------------- '
' FileExists
' --------------------------------------------- '

'@testmethod FileSystem.FileExists
Private Sub FileExists_EmptyString()
    This.Assert.AreEqual This.ScrFileSystem.FileExists(vbNullString), This.VbaFileSystem.FileExists(vbNullString)
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_True()
    This.Assert.AreEqual This.ScrFileSystem.FileExists(This.TestFilePath), This.VbaFileSystem.FileExists(This.TestFilePath)
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_False()
    This.Assert.AreEqual This.ScrFileSystem.FileExists("C:\Users\JohnDoe\Documents\Hello World.txt"), This.VbaFileSystem.FileExists("C:\Users\JohnDoe\Documents\Hello World.txt")
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_RelativePath()
    Dim test_RelPath As String
    test_RelPath = This.ScrFileSystem.BuildPath("..", This.ScrFileSystem.GetBaseName(This.TestFilePath))
    This.Assert.AreEqual This.ScrFileSystem.FileExists(test_RelPath), This.VbaFileSystem.FileExists(test_RelPath)
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_InvalidPath()
    This.Assert.AreEqual This.ScrFileSystem.FileExists("Hello World.txt"), This.VbaFileSystem.FileExists("Hello World.txt")
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_EndWithFileSeparator()
    This.Assert.AreEqual This.ScrFileSystem.FileExists(This.TestFilePath & Application.PathSeparator), This.VbaFileSystem.FileExists(This.TestFilePath & Application.PathSeparator)
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_FileSeparatorForwardslash()
    This.Assert.AreEqual This.ScrFileSystem.FileExists(VBA.Replace(This.TestFilePath, "\", "/")), This.VbaFileSystem.FileExists(VBA.Replace(This.TestFilePath, "\", "/"))
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_NoFileExtension()
    Dim test_Path As String
    test_Path = VBA.Replace(This.TestFilePath, "." & This.VbaFileSystem.GetExtensionName(This.TestFilePath), vbNullString)
    This.Assert.AreEqual This.ScrFileSystem.FileExists(test_Path), This.VbaFileSystem.FileExists(test_Path)
End Sub
'@testmethod FileSystem.FileExists
Private Sub FileExists_Folder()
    This.Assert.AreEqual This.ScrFileSystem.FileExists(This.TestFolderPath), This.VbaFileSystem.FileExists(This.TestFolderPath)
End Sub

' --------------------------------------------- '
' FolderExists
' --------------------------------------------- '

'@testmethod FileSystem.FolderExists
Private Sub FolderExists_EmptyString()
    This.Assert.AreEqual This.ScrFileSystem.FolderExists(vbNullString), This.VbaFileSystem.FolderExists(vbNullString)
End Sub
'@testmethod FileSystem.FolderExists
Private Sub FolderExists_True()
    This.Assert.AreEqual This.ScrFileSystem.FolderExists(This.TestFolderPath), This.VbaFileSystem.FolderExists(This.TestFolderPath)
End Sub
'@testmethod FileSystem.FolderExists
Private Sub FolderExists_False()
    This.Assert.AreEqual This.ScrFileSystem.FolderExists("C:\Users\JohnDoe\Documents"), This.VbaFileSystem.FolderExists("C:\Users\JohnDoe\Documents")
End Sub
'@testmethod FileSystem.FolderExists
Private Sub FolderExists_RelativePath()
    Dim test_RelPath As String
    test_RelPath = This.ScrFileSystem.BuildPath("..", This.ScrFileSystem.GetBaseName(This.TestFolderPath))
    This.Assert.AreEqual This.ScrFileSystem.FolderExists(test_RelPath), This.VbaFileSystem.FolderExists(test_RelPath)
End Sub
'@testmethod FileSystem.FolderExists
Private Sub FolderExists_InvalidPath()
    This.Assert.AreEqual This.ScrFileSystem.FolderExists("Hello World"), This.VbaFileSystem.FolderExists("Hello World")
End Sub
'@testmethod FileSystem.FolderExists
Private Sub FolderExists_EndWithFileSeparator()
    This.Assert.AreEqual This.ScrFileSystem.FolderExists(This.TestFolderPath & Application.PathSeparator), This.VbaFileSystem.FolderExists(This.TestFolderPath & Application.PathSeparator)
End Sub
'@testmethod FileSystem.FolderExists
Private Sub FolderExists_FileSeparatorForwardslash()
    This.Assert.AreEqual This.ScrFileSystem.FolderExists(VBA.Replace(This.TestFolderPath, "\", "/")), This.VbaFileSystem.FolderExists(VBA.Replace(This.TestFolderPath, "\", "/"))
End Sub
'@testmethod FileSystem.FolderExists
Private Sub FolderExists_File()
    This.Assert.AreEqual This.ScrFileSystem.FolderExists(This.TestFilePath), This.VbaFileSystem.FolderExists(This.TestFilePath)
End Sub

' --------------------------------------------- '
' GetBaseName
' --------------------------------------------- '

'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_EmptyString()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName(vbNullString), This.VbaFileSystem.GetBaseName(vbNullString)
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_NoFileSeparator()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName("Hello World"), This.VbaFileSystem.GetBaseName("Hello World")
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_RelativePath()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName("..\Documents\HelloWorld.txt"), This.VbaFileSystem.GetBaseName("..\Documents\HelloWorld.txt")
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_FileSeparatorBackslash()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName("C:\Users\JohnDoe\Documents\HelloWorld.txt"), This.VbaFileSystem.GetBaseName("C:\Users\JohnDoe\Documents\HelloWorld.txt")
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_FileSeparatorForwardslash()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName("https://www.google.co.nz/HelloWorld.txt"), This.VbaFileSystem.GetBaseName("https://www.google.co.nz/HelloWorld.txt")
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_FileSeparatorMixed()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName("C:\Users\JohnDoe\Documents/HelloWorld.txt"), This.VbaFileSystem.GetBaseName("C:\Users\JohnDoe\Documents/HelloWorld.txt")
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_EndWithFileSeparator()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName("C:\Users\JohnDoe\Documents\"), This.VbaFileSystem.GetBaseName("C:\Users\JohnDoe\Documents\")
End Sub
'@testmethod FileSystem.GetBaseName
Private Sub GetBaseName_DriveSpecOnly()
    This.Assert.AreEqual This.ScrFileSystem.GetBaseName("C:\"), This.VbaFileSystem.GetBaseName("C:\")
End Sub

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
'@testmethod FileSystem.SpeedTest
Private Sub speedtest_FileExists()
    Dim test_Temp As Boolean
    Dim test_Long As Long
    Dim test_StartTime As Date
    Dim test_FinishTime As Date
    Dim test_VbaMS As Double
    Dim test_ScrMs As Double
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 25000
        test_Temp = This.ScrFileSystem.FileExists("C:\Users\JohnDoe\Documents\Hello World.txt") ' False
        test_Temp = This.ScrFileSystem.FileExists(This.TestFilePath) ' True
    Next test_Long
    test_FinishTime = VBA.Date + CDate(VBA.Timer / 86400)
    test_ScrMs = VBA.Round((test_FinishTime - test_StartTime) * 86400 * 1000, 4)
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 25000
        test_Temp = This.VbaFileSystem.FileExists("C:\Users\JohnDoe\Documents\Hello World.txt") ' False
        test_Temp = This.VbaFileSystem.FileExists(This.TestFilePath) ' True
    Next test_Long
    test_FinishTime = VBA.Date + CDate(VBA.Timer / 86400)
    test_VbaMS = VBA.Round((test_FinishTime - test_StartTime) * 86400 * 1000, 4)
    
    This.Assert.Inconclusive "SCR=" & test_ScrMs & "ms | VBA=" & test_VbaMS & "ms | " & VBA.IIf(test_VbaMS > test_ScrMs, "Scripting", "VBA") & " is " & VBA.Round(VBA.IIf(test_VbaMS > test_ScrMs, test_VbaMS / test_ScrMs, test_ScrMs / test_VbaMS), 4) & " times faster."
End Sub
'@testmethod FileSystem.SpeedTest
Private Sub speedtest_FolderExists()
    Dim test_Temp As Boolean
    Dim test_Long As Long
    Dim test_StartTime As Date
    Dim test_FinishTime As Date
    Dim test_VbaMS As Double
    Dim test_ScrMs As Double
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 25000
        test_Temp = This.ScrFileSystem.FolderExists("C:\Users\JohnDoe\Documents") ' False
        test_Temp = This.ScrFileSystem.FolderExists(This.TestFolderPath) ' True
    Next test_Long
    test_FinishTime = VBA.Date + CDate(VBA.Timer / 86400)
    test_ScrMs = VBA.Round((test_FinishTime - test_StartTime) * 86400 * 1000, 4)
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 25000
        test_Temp = This.VbaFileSystem.FolderExists("C:\Users\JohnDoe\Documents") ' False
        test_Temp = This.VbaFileSystem.FolderExists(This.TestFolderPath) ' True
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
        .TestFolderPath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "TestsFolder")
        .TestFilePath = .ScrFileSystem.BuildPath(.TestFolderPath, "TestFile.txt")
        If Not .ScrFileSystem.FolderExists(.TestFolderPath) Then .ScrFileSystem.CreateFolder .TestFolderPath
        If Not .ScrFileSystem.FileExists(.TestFilePath) Then .ScrFileSystem.CreateTextFile .TestFilePath
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    With This
        If .ScrFileSystem.FolderExists(.TestFolderPath) Then .ScrFileSystem.DeleteFolder (.TestFolderPath)
        If .ScrFileSystem.FileExists(.TestFilePath) Then .ScrFileSystem.DeleteFile (.TestFilePath)
        Set .Assert = Nothing
        Set .Fakes = Nothing
        Set .ScrFileSystem = Nothing
        Set .VbaFileSystem = Nothing
    End With
End Sub
