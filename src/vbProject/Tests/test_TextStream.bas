Attribute VB_Name = "test_TextStream"
'<RC OBFUSCATION=DEVELOPMENT MODULE>
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

Private Type TResult
    Expected As Variant
    Actual As Variant
End Type

Private Type TTest
    Assert As Object
    Fakes As Object
    ScrFileSystem As Object
    VbaFileSystem As FileSystemObject
    unicode_ScrFilePath As String
    unicode_VbaFilePath As String
    ascii_ScrFilePath As String
    ascii_VbaFilePath As String
End Type

Private This As TTest

' ============================================= '
' Test Methods
' ============================================= '

''
' Test that VBA TextStream writes the same bytes as Scripting TextStream when using Unicode (TristateTrue) format.
'
'@testmethod Scripting.TextStream
''
Private Sub unicode_Write()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Write(TristateTrue)
'---ASSERT---
    This.Assert.AreEqual test_Result.Expected, test_Result.Actual, "File contents do not match when writing Unicode bytes to file."
End Sub


''
' Test that VBA TextStream writes the same bytes as Scripting TextStream when using ASCII (TristateFalse) format.
'
'@testmethod Scripting.TextStream
''
Private Sub ascii_Write()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Write(TristateFalse)
'---ASSERT---
    This.Assert.AreEqual test_Result.Expected, test_Result.Actual, "File contents do not match when writing ASCII bytes to file."
End Sub

''
' Test that VBA TextStream appends the same bytes as Scripting TextStream when using Unicode(TristateTrue) format.
'
'@testmethod Scripting.TextStream
''
Private Sub unicode_Append()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Append(TristateTrue)
'---ASSERT---
    This.Assert.AreEqual test_Result.Expected, test_Result.Actual, "File contents do not match when appending Unicode bytes to file." & vbNewLine & _
                                                                   "NOTE: If unicode_Write failed, this test will probably fail."
End Sub

''
' Test that VBA TextStream appends the same bytes as Scripting TextStream when using ASCII (TristateFalse) format.
'
'@testmethod Scripting.TextStream
''
Private Sub ascii_Append()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Append(TristateFalse)
'---ASSERT---
    This.Assert.AreEqual test_Result.Expected, test_Result.Actual, "File contents do not match when appending ASCII bytes to file." & vbNewLine & _
                                                                   "NOTE: If ascii_Write failed, this test will probably fail."
End Sub

''
' Read characters encoded using Unicode (TristateTrue) with VBA TextStream and Scripting TextStream and compare results.
'
'@testmethod Scripting.TextStream
''
Private Sub unicode_ReadChr()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_ReadChr(TristateTrue, TristateTrue)
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "Reading contents by character produced differing results." & vbNewLine & _
                                                                         "NOTE: both 'Skip' and 'SkipLine' are used here, ensure these tests are passing in the first instance"
End Sub
''
' Read characters encoded using ASCII (TristateFalse) with VBA TextStream and Scripting TextStream and compare results.
'
'@testmethod Scripting.TextStream
''
Private Sub ascii_ReadChr()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_ReadChr(TristateFalse, TristateFalse)
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "Reading contents by character produced differing results." & vbNewLine & _
                                                                         "NOTE: both 'Skip' and 'SkipLine' are used here, ensure these tests are passing in the first instance"
End Sub


' ============================================= '
' Private Methods
' ============================================= '

''
' Write contents to file in `WriteFormat` using TextStream methods:
'  - WriteLine
'  - Write
'  - WriteBlankLines
'
' @param {Tristate} WriteFormat | Byte format to write file contents with.
''
Private Function private_Write(ByVal WriteFormat As Tristate) As TResult
    ' Variables.
    Dim test_ScrStream As Object
    Dim test_VbaStream As TextStream
    Dim test_ScrFilePath As String
    Dim test_VbaFilePath As String
    
'---ARRANGE---
    ' Create files.
    Select Case WriteFormat
    Case Tristate.TristateTrue ' Unicode
        test_ScrFilePath = This.unicode_ScrFilePath
        test_VbaFilePath = This.unicode_VbaFilePath
        Set test_ScrStream = This.ScrFileSystem.CreateTextFile(test_ScrFilePath, True, True)
        Set test_VbaStream = This.VbaFileSystem.CreateTextFile(test_VbaFilePath, True, True)
    Case Tristate.TristateFalse ' ASCII
        test_ScrFilePath = This.ascii_ScrFilePath
        test_VbaFilePath = This.ascii_VbaFilePath
        Set test_ScrStream = This.ScrFileSystem.CreateTextFile(test_ScrFilePath, True, False)
        Set test_VbaStream = This.VbaFileSystem.CreateTextFile(test_VbaFilePath, True, False)
    Case Else
        ' TODO - currently not supported.
    End Select
    
'---ACT---
    ' Write the same contents to file using Scripting and VisualBasic.
    With test_ScrStream
        .WriteLine "Hello World"
        .WriteLine "Hello World(2)"
        .WriteLine "Hello World(3)"
        .Write "Hello World"
        .Write "Hello World"
        .WriteBlankLines 4
        .Close
    End With
    With test_VbaStream
        .WriteLine "Hello World"
        .WriteLine "Hello World(2)"
        .WriteLine "Hello World(3)"
        .WriteStr "Hello World"
        .WriteStr "Hello World"
        .WriteBlankLines 4
        .CloseFile
    End With
    
'---ASSERT---
    ' Read contents of both files using Scripting.
    private_Write.Expected = This.ScrFileSystem.OpenTextFile(test_ScrFilePath, ForReading, False, WriteFormat).ReadAll
    private_Write.Actual = This.ScrFileSystem.OpenTextFile(test_VbaFilePath, ForReading, False, WriteFormat).ReadAll
End Function

''
' Create a file in `WriteFormat`, write a few lines then close. Open in Append and append content
' using `WriteFormat`.
'
' Used by tests:
'  - unicode_Append
'  - ascii_Append
'
' @param {Tristate} WriteFormat | Byte format to write/append file contents with.
''
Private Function private_Append(ByVal WriteFormat As Tristate) As TResult
    ' Variables.
    Dim test_ScrStream As Object
    Dim test_VbaStream As TextStream
    Dim test_ScrFilePath As String
    Dim test_VbaFilePath As String
    
'---ARRANGE---
    ' Create files & write a few lines.
    Select Case WriteFormat
    Case Tristate.TristateTrue ' Unicode
        test_ScrFilePath = This.unicode_ScrFilePath
        test_VbaFilePath = This.unicode_VbaFilePath
        Set test_ScrStream = This.ScrFileSystem.CreateTextFile(test_ScrFilePath, True, True)
        Set test_VbaStream = This.VbaFileSystem.CreateTextFile(test_VbaFilePath, True, True)
    Case Tristate.TristateFalse ' ASCII
        test_ScrFilePath = This.ascii_ScrFilePath
        test_VbaFilePath = This.ascii_VbaFilePath
        Set test_ScrStream = This.ScrFileSystem.CreateTextFile(test_ScrFilePath, True, False)
        Set test_VbaStream = This.VbaFileSystem.CreateTextFile(test_VbaFilePath, True, False)
    Case Else
        ' TODO - currently not supported.
    End Select
    test_ScrStream.WriteLine "This is the first line in the file, which will have content appended."
    test_VbaStream.WriteLine "This is the first line in the file, which will have content appended."
    test_ScrStream.Close
    test_VbaStream.CloseFile
    
    ' Open files for appending.
    Set test_ScrStream = This.ScrFileSystem.OpenTextFile(test_ScrFilePath, ForAppending, False, WriteFormat)
    Set test_VbaStream = This.VbaFileSystem.OpenTextFile(test_VbaFilePath, ForAppending, False, WriteFormat)

'---ACT---
    ' Append the same contents to file using Scripting and VisualBasic.
    With test_ScrStream
        .WriteLine vbNullString
        .Write vbNullString
        .WriteLine "AppendLine1"
        .WriteLine "AppendLine2"
        .WriteLine "AppendLine3"
        .Write "AppendLine4"
        .Write "...Continued line4"
        .WriteBlankLines 4
        .Close
    End With
    With test_VbaStream
        .WriteLine vbNullString
        .WriteStr vbNullString
        .WriteLine "AppendLine1"
        .WriteLine "AppendLine2"
        .WriteLine "AppendLine3"
        .WriteStr "AppendLine4"
        .WriteStr "...Continued line4"
        .WriteBlankLines 4
        .CloseFile
    End With
    
'---ASSERT---
    ' Read contents of both files using Scripting.
    private_Append.Expected = This.ScrFileSystem.OpenTextFile(test_ScrFilePath, ForReading, False, WriteFormat).ReadAll
    private_Append.Actual = This.ScrFileSystem.OpenTextFile(test_VbaFilePath, ForReading, False, WriteFormat).ReadAll
End Function

''
' Create file in `WriteFormat` then read using `TextStream.Read` method (specifiying number of characters) in `ReadFormat`.
'
' Used by tests:
'  - unicode_ReadChr
'  - ascii_ReadChr
''
Private Function private_ReadChr(ByVal WriteFormat As Tristate, ByVal ReadFormat As Tristate) As TResult
    Dim test_ScrStream As Object
    Dim test_VbaStream As TextStream
    Dim test_ScrChar(1 To 5) As String
    Dim test_VbaChar(1 To 5) As String
    Dim test_TargetFilePath As String
    
'---ARRANGE---
    ' Create files & write a few lines.
    Select Case WriteFormat
    Case Tristate.TristateFalse ' ASCII
        test_TargetFilePath = This.ascii_ScrFilePath
        Set test_ScrStream = This.ScrFileSystem.CreateTextFile(test_TargetFilePath, True, False)
    Case Tristate.TristateTrue ' Unicode.
        test_TargetFilePath = This.unicode_ScrFilePath
        Set test_ScrStream = This.ScrFileSystem.CreateTextFile(test_TargetFilePath, True, True)
    Case Else
        ' TODO - currently not supported.
    End Select
    With test_ScrStream
        .WriteLine vbNullString
        .Write vbNullString
        .WriteLine "SampleLineOne"
        .WriteLine "SampleLineTwo"
        .WriteLine "SampleLineThree"
        .WriteBlankLines 4
    End With
    
'---ACT---
    ' Read random characters.
    Set test_ScrStream = This.ScrFileSystem.OpenTextFile(test_TargetFilePath, ForReading, False, ReadFormat)
    With test_ScrStream
        test_ScrChar(1) = .Read(1)
        test_ScrChar(2) = .Read(5)
        .Skip 15
        test_ScrChar(3) = .Read(2)
        test_ScrChar(4) = .Read(4)
        .SkipLine
        test_ScrChar(5) = .Read(1)
        .Close
    End With
    Set test_VbaStream = This.VbaFileSystem.OpenTextFile(test_TargetFilePath, ForReading, False, ReadFormat)
    With test_VbaStream
        test_VbaChar(1) = .Read(1)
        test_VbaChar(2) = .Read(5)
        .Skip 15
        test_VbaChar(3) = .Read(2)
        test_VbaChar(4) = .Read(4)
        .SkipLine
        test_VbaChar(5) = .Read(1)
        .CloseFile
    End With
    
'---ASSERT---
    private_ReadChr.Expected = test_ScrChar
    private_ReadChr.Actual = test_VbaChar
End Function

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Dim test_Item As Variant
    With This
        For Each test_Item In Array(.unicode_ScrFilePath, .unicode_VbaFilePath, .ascii_ScrFilePath, .ascii_VbaFilePath)
            If .ScrFileSystem.FileExists(test_Item) Then .ScrFileSystem.DeleteFile test_Item, True
        Next test_Item
    End With
End Sub

'@ModuleInitialize
Private Sub ModuleInitialize()
    With This
        Set .Assert = CreateObject("Rubberduck.AssertClass")
        Set .Fakes = CreateObject("Rubberduck.FakesProvider")
        Set .ScrFileSystem = CreateObject("Scripting.FileSystemObject")
        Set .VbaFileSystem = New FileSystemObject
        .unicode_ScrFilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_unicode_Scripting.txt")
        .unicode_VbaFilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_unicode_VisualBasic.txt")
        .ascii_ScrFilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_ascii_Scripting.txt")
        .ascii_VbaFilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_ascii_VisualBasic.txt")
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


