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

Private Type TTest
    Assert As Object
    Fakes As Object
End Type

Private This As TTest

' ============================================= '
' Test Methods
' ============================================= '

' TODO - Convert this code to actual tests, not the sudo-tests they currently are.
Private Sub Testspeedactual()
    Dim scr_FilePath As String
    Dim vba_FilePath As String
    Dim test_Long As Long
    Dim scr_FSO As Object
    Dim scr_TextStream As Object
    Dim vba_FSO As FileSystemObject
    Dim vba_TextStream As TextStream
    Dim test_Unicode As Boolean
    Dim test_String As String
    
    scr_FilePath = "C:\Users\AndrewPullon\Documents\repos\RadiusCore\radius-excel\scr.txt"
    vba_FilePath = "C:\Users\AndrewPullon\Documents\repos\RadiusCore\radius-excel\vba.txt"
    Set scr_FSO = CreateObject("Scripting.FileSystemObject")
    Set vba_FSO = New FileSystemObject
    
    test_Unicode = False
    
    Set scr_TextStream = scr_FSO.CreateTextFile(scr_FilePath, True, test_Unicode)
    Set vba_TextStream = vba_FSO.CreateTextFile(vba_FilePath, True, test_Unicode)
    
    Debug.Print "Scripting START: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
    'For test_Long = 1 To 20000
    '    scr_TextStream.WriteLine "Hello World"
    'Next test_Long
    scr_TextStream.WriteBlankLines 20000
    scr_TextStream.Close
    Debug.Print "Scripting   END: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
    
    Debug.Print "VBA       START: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
    'For test_Long = 1 To 20000
    '    vba_TextStream.WriteLine "Hello World"
    'Next test_Long
    vba_TextStream.WriteBlankLines 20000
    vba_TextStream.CloseFile
    Debug.Print "VBA         END: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
    
    Debug.Print "Scripting START: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
        Set scr_TextStream = scr_FSO.OpenTextFile(scr_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
        test_String = scr_TextStream.ReadAll
    Debug.Print "Scripting   END: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
    Debug.Print "VBA       START: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
        Set vba_TextStream = vba_FSO.OpenTextFile(vba_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
        test_String = vba_TextStream.ReadAll
    Debug.Print "VBA         END: " & VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2)
End Sub

Private Sub testspeed()
    Dim scr_FilePath As String
    Dim vba_FilePath As String
    Dim scr_FSO As Object
    Dim scr_TextStream As Object
    Dim vba_FSO As FileSystemObject
    Dim vba_TextStream As TextStream
    Dim vba_Str As String
    Dim scr_Str As String
    Dim test_Long As Long
    Dim test_Item As Variant
    Dim test_Unicode As Boolean
    Dim test_Loop As Boolean
    
    test_Loop = True
    scr_FilePath = "C:\Users\AndrewPullon\Documents\repos\RadiusCore\radius-excel\scr.txt"
    vba_FilePath = "C:\Users\AndrewPullon\Documents\repos\RadiusCore\radius-excel\vba.txt"
    Set scr_FSO = CreateObject("Scripting.FileSystemObject")
    Set vba_FSO = New FileSystemObject
    
    test_Unicode = True
    
    ' Create files.
    Set scr_TextStream = scr_FSO.CreateTextFile(scr_FilePath, True, test_Unicode)
    With scr_TextStream
        .WriteLine "Hello World"
        .WriteLine "Hello World(2)"
        .WriteLine "Hello World(3)"
        .Write "Heloow!!!!"
        .Write "Heloow!!!!"
        .WriteBlankLines 4
        .Close
    End With
    Set vba_TextStream = vba_FSO.CreateTextFile(vba_FilePath, True, test_Unicode)
    With vba_TextStream
        .WriteLine "Hello World"
        .WriteLine "Hello World(2)"
        .WriteLine "Hello World(3)"
        .WriteStr "Heloow!!!!"
        .WriteStr "Heloow!!!!"
        .WriteBlankLines 4
        .CloseFile
    End With

testloop:
    ' Verify contents are the same, using Scripting..
    Set scr_TextStream = scr_FSO.OpenTextFile(scr_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
    scr_Str = scr_TextStream.ReadAll
    scr_TextStream.Close
    Set scr_TextStream = scr_FSO.OpenTextFile(vba_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
    vba_Str = scr_TextStream.ReadAll
    scr_TextStream.Close
    Debug.Print scr_Str = vba_Str
    
    ' Verify contents are the same, using VBA.
    Set vba_TextStream = vba_FSO.OpenTextFile(scr_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
    scr_Str = vba_TextStream.ReadAll
    vba_TextStream.CloseFile
    Set vba_TextStream = vba_FSO.OpenTextFile(vba_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
    vba_Str = vba_TextStream.ReadAll
    vba_TextStream.CloseFile
    Debug.Print scr_Str = vba_Str
    
    ' Compare read of same file, different TextStreams.
    For Each test_Item In Array(scr_FilePath, vba_FilePath)
        Set scr_TextStream = scr_FSO.OpenTextFile(test_Item, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
        scr_Str = scr_TextStream.ReadAll
        scr_TextStream.Close
        Set vba_TextStream = vba_FSO.OpenTextFile(test_Item, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
        vba_Str = vba_TextStream.ReadAll
        vba_TextStream.CloseFile
        Debug.Print scr_Str = vba_Str
    Next test_Item
    
    If test_Loop Then
        ' Append content to file, then perform same checks.
        Set scr_TextStream = scr_FSO.OpenTextFile(scr_FilePath, ForAppending, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
        With scr_TextStream
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
        Set vba_TextStream = vba_FSO.OpenTextFile(vba_FilePath, ForAppending, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
        With vba_TextStream
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
        test_Loop = False
        GoTo testloop
    End If
    
    ' Test properties
    Set scr_TextStream = scr_FSO.OpenTextFile(scr_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
    scr_Str = scr_TextStream.Read(11)
    'scr_Str = scr_TextStream.ReadAll
    
    Set vba_TextStream = vba_FSO.OpenTextFile(vba_FilePath, ForReading, False, VBA.IIf(test_Unicode, TristateTrue, TristateFalse))
    vba_Str = vba_TextStream.Read(11)
    'vba_Str = vba_TextStream.ReadAll
    
    Debug.Print scr_Str = vba_Str
    Debug.Print scr_TextStream.Line = vba_TextStream.Line
    Debug.Print scr_TextStream.Column = vba_TextStream.Column
    
    scr_TextStream.Close
    vba_TextStream.CloseFile
    
    'Set scr_TextStream = scr_FSO.OpenTextFile("C:\Users\AndrewPullon\Documents\repos\RadiusCore\radius-excel\logs\RadiusCore.txt", ForReading, False, TristateFalse)
    'scr_TextStream.Skip 200
    'scr_Str = scr_TextStream.ReadLine
    'scr_TextStream.Close
    'Set vba_TextStream = vba_FSO.OpenTextFile("C:\Users\AndrewPullon\Documents\repos\RadiusCore\radius-excel\logs\RadiusCore.txt", ForReading, False, TristateFalse)
    'vba_TextStream.Skip 200
    'vba_Str = vba_TextStream.ReadLine
    'vba_TextStream.CloseFile
    'Debug.Print scr_Str = vba_Str
End Sub

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set This.Assert = CreateObject("Rubberduck.AssertClass")
    Set This.Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set This.Assert = Nothing
    Set This.Fakes = Nothing
End Sub


