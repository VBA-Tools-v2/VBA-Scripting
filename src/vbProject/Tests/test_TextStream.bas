Attribute VB_Name = "test_TextStream"
''
' VBA-Git Annotations
' https://github.com/VBA-Tools-v2/VBA-Git | https://radiuscore.co.nz
'
' @developmentmodule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
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
    ScrFileSystem As Object             ' Scripting.FileSystemObject
    VbaFileSystem As FileSystemObject   ' VBA.FileSystemObject
    scr_FilePath As String              ' A file created using Scripting.TextStream.
    vba_FilePath As String              ' A file created using VBA.TextStream.
    unicode_FilePath As String          ' A file encoded in Unicode format.
    ascii_FilePath As String            ' A file encoded in ASCII format.
    generic_FilePath As String          ' A file of no specific format/origin.
End Type

Private This As TTest

' ============================================= '
' Test Methods
' ============================================= '

' --------------------------------------------- '
' Write Tests
' --------------------------------------------- '

''
' Test that VBA TextStream writes the same bytes as Scripting TextStream when using Unicode (TristateTrue) format.
'
'@testmethod Scripting.TextStream.Unicode
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
'@testmethod Scripting.TextStream.ASCII
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
'@testmethod Scripting.TextStream.Unicode
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
'@testmethod Scripting.TextStream.ASCII
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

' --------------------------------------------- '
' Read Tests
' --------------------------------------------- '

''
' Read characters encoded using Unicode (TristateTrue) with VBA TextStream and Scripting TextStream and compare results.
'
'@testmethod Scripting.TextStream.Unicode
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
'@testmethod Scripting.TextStream.ASCII
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

' --------------------------------------------- '
' Property Tests
' --------------------------------------------- '
' Read file encoded using Unicode (TristateTrue) or ASCII (TristateFalse) with VBA TextStream and Scripting TextStream and compare properties.

'@testmethod Scripting.TextStream.Unicode
Private Sub unicode_property_Column()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateTrue, TristateTrue, "Column")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "Column property is not correct when reading Unicode file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub
'@testmethod Scripting.TextStream.ASCII
Private Sub ascii_property_Column()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateFalse, TristateFalse, "Column")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "Column property is not correct when reading ASCII file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub
'@testmethod Scripting.TextStream.Unicode
Private Sub unicode_property_Line()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateTrue, TristateTrue, "Line")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "Line property is not correct when reading Unicode file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub
'@testmethod Scripting.TextStream.ASCII
Private Sub ascii_property_Line()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateFalse, TristateFalse, "Line")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "Line property is not correct when reading ASCII file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub
'@testmethod Scripting.TextStream.Unicode
Private Sub unicode_property_AtEndOfLine()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateTrue, TristateTrue, "AtEndOfLine")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "AtEndOfLine property is not correct when reading Unicode file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub
'@testmethod Scripting.TextStream.ASCII
Private Sub ascii_property_AtEndOfLine()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateFalse, TristateFalse, "AtEndOfLine")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "AtEndOfLine property is not correct when reading ASCII file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub
'@testmethod Scripting.TextStream.Unicode
Private Sub unicode_property_AtEndOfStream()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateTrue, TristateTrue, "AtEndOfStream")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "AtEndOfStream property is not correct when reading Unicode file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub
'@testmethod Scripting.TextStream.ASCII
Private Sub ascii_property_AtEndOfStream()
'---ARRANGE---
    Dim test_Result As TResult
'---ACT---
    test_Result = private_Properties(TristateFalse, TristateFalse, "AtEndOfStream")
'---ASSERT---
    This.Assert.SequenceEquals test_Result.Expected, test_Result.Actual, "AtEndOfStream property is not correct when reading ASCII file." & vbNewLine & _
                                                                         "NOTE: Ensure all 'Read' tests are passing if this test is failing."
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Function private_Properties(ByVal FileFormat As Tristate, ByVal ReadFormat As Tristate, ByVal TargetProperty As String) As TResult
    ' Variables.
    Dim test_ScrStream As Object
    Dim test_VbaStream As TextStream
    Dim test_ScrProp(1 To 5) As Variant
    Dim test_VbaProp(1 To 5) As Variant
    
'---ARRANGE---
    ' Create files on disk with contents.
    private_CreateDummyFile This.scr_FilePath, FileFormat, True
    private_CreateDummyFile This.vba_FilePath, FileFormat, True
    ' Open files for reading.
    Set test_ScrStream = This.ScrFileSystem.OpenTextFile(This.scr_FilePath, ForReading, False, ReadFormat)
    Set test_VbaStream = This.VbaFileSystem.OpenTextFile(This.vba_FilePath, ForReading, False, ReadFormat)
'---ACT---
    ' Read contents, storing properties
    With test_ScrStream
        .Read 10
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_ScrProp(1) = .AtEndOfLine
        Case "AtEndOfStream"
            test_ScrProp(1) = .AtEndOfStream
        Case "Column"
            test_ScrProp(1) = .Column
        Case "Line"
            test_ScrProp(1) = .Line
        End Select
        
        .Read 2
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_ScrProp(2) = .AtEndOfLine
        Case "AtEndOfStream"
            test_ScrProp(2) = .AtEndOfStream
        Case "Column"
            test_ScrProp(2) = .Column
        Case "Line"
            test_ScrProp(2) = .Line
        End Select
        
        .ReadLine
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_ScrProp(3) = .AtEndOfLine
        Case "AtEndOfStream"
            test_ScrProp(3) = .AtEndOfStream
        Case "Column"
            test_ScrProp(3) = .Column
        Case "Line"
            test_ScrProp(3) = .Line
        End Select
        
        .Read 4
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_ScrProp(4) = .AtEndOfLine
        Case "AtEndOfStream"
            test_ScrProp(4) = .AtEndOfStream
        Case "Column"
            test_ScrProp(4) = .Column
        Case "Line"
            test_ScrProp(4) = .Line
        End Select
        
        .ReadAll
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_ScrProp(5) = .AtEndOfLine
        Case "AtEndOfStream"
            test_ScrProp(5) = .AtEndOfStream
        Case "Column"
            test_ScrProp(5) = .Column
        Case "Line"
            test_ScrProp(5) = .Line
        End Select
        .Close
    End With
    With test_VbaStream
        .Read 10
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_VbaProp(1) = .AtEndOfLine
        Case "AtEndOfStream"
            test_VbaProp(1) = .AtEndOfStream
        Case "Column"
            test_VbaProp(1) = .Column
        Case "Line"
            test_VbaProp(1) = .Line
        End Select
        
        .Read 2
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_VbaProp(2) = .AtEndOfLine
        Case "AtEndOfStream"
            test_VbaProp(2) = .AtEndOfStream
        Case "Column"
            test_VbaProp(2) = .Column
        Case "Line"
            test_VbaProp(2) = .Line
        End Select
        
        .ReadLine
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_VbaProp(3) = .AtEndOfLine
        Case "AtEndOfStream"
            test_VbaProp(3) = .AtEndOfStream
        Case "Column"
            test_VbaProp(3) = .Column
        Case "Line"
            test_VbaProp(3) = .Line
        End Select
        
        .Read 4
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_VbaProp(4) = .AtEndOfLine
        Case "AtEndOfStream"
            test_VbaProp(4) = .AtEndOfStream
        Case "Column"
            test_VbaProp(4) = .Column
        Case "Line"
            test_VbaProp(4) = .Line
        End Select
        
        .ReadAll
        Select Case TargetProperty
        Case "AtEndOfLine"
            test_VbaProp(5) = .AtEndOfLine
        Case "AtEndOfStream"
            test_VbaProp(5) = .AtEndOfStream
        Case "Column"
            test_VbaProp(5) = .Column
        Case "Line"
            test_VbaProp(5) = .Line
        End Select
        .CloseFile
    End With
'---ASSERT---
    private_Properties.Expected = test_ScrProp
    private_Properties.Actual = test_VbaProp
End Function

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
    
'---ARRANGE---
    ' Create files.
    Set test_ScrStream = This.ScrFileSystem.CreateTextFile(This.scr_FilePath, True, WriteFormat = TristateTrue)
    Set test_VbaStream = This.VbaFileSystem.CreateTextFile(This.vba_FilePath, True, WriteFormat = TristateTrue)
    
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
    private_Write.Expected = This.ScrFileSystem.OpenTextFile(This.scr_FilePath, ForReading, False, WriteFormat).ReadAll
    private_Write.Actual = This.ScrFileSystem.OpenTextFile(This.vba_FilePath, ForReading, False, WriteFormat).ReadAll
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
    
'---ARRANGE---
    ' Create dummy file with contents.
    private_CreateDummyFile This.scr_FilePath, WriteFormat, True
    private_CreateDummyFile This.vba_FilePath, WriteFormat, True
    
    ' Open files for appending.
    Set test_ScrStream = This.ScrFileSystem.OpenTextFile(This.scr_FilePath, ForAppending, False, WriteFormat)
    Set test_VbaStream = This.VbaFileSystem.OpenTextFile(This.vba_FilePath, ForAppending, False, WriteFormat)

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
    private_Append.Expected = This.ScrFileSystem.OpenTextFile(This.scr_FilePath, ForReading, False, WriteFormat).ReadAll
    private_Append.Actual = This.ScrFileSystem.OpenTextFile(This.vba_FilePath, ForReading, False, WriteFormat).ReadAll
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
    ' Create dummy file with contents.
    private_CreateDummyFile This.generic_FilePath, WriteFormat, True
    
'---ACT---
    ' Read random characters.
    Set test_ScrStream = This.ScrFileSystem.OpenTextFile(This.generic_FilePath, ForReading, False, ReadFormat)
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
    Set test_VbaStream = This.VbaFileSystem.OpenTextFile(This.generic_FilePath, ForReading, False, ReadFormat)
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

''
' Create a file on disk using Scripting library, with given `WriteFormat`.
''
Private Sub private_CreateDummyFile(ByVal FilePath As String, ByVal WriteFormat As Tristate, ByVal IncludeContents As Boolean)
    Dim test_Stream As Object
    Set test_Stream = This.ScrFileSystem.CreateTextFile(FilePath, True, WriteFormat = TristateTrue)
    
    If IncludeContents Then
        With test_Stream
            .WriteLine "Hello World"
            .WriteLine "Hello World(2)"
            .WriteLine "Hello World(3)"
            .Write "Hello World"
            .Write "Hello World"
            .WriteBlankLines 4
        End With
    End If
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
    Dim test_Item As Variant
    With This
        For Each test_Item In Array(.ascii_FilePath, .unicode_FilePath, .scr_FilePath, .vba_FilePath, .generic_FilePath)
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
        .ascii_FilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_Ascii.txt")
        .unicode_FilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_Unicode.txt")
        .scr_FilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_Scripting.txt")
        .vba_FilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_VisualBasic.txt")
        .generic_FilePath = .ScrFileSystem.BuildPath(ThisWorkbook.Path, "test_Generic.txt")
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
