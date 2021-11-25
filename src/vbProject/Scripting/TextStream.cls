VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' TextStream | STATUS: In Development.
' (c) RadiusCore Ltd - https://radiuscore.co.nz/
'
' Facilitates sequential access to file.
'
' Native VBA implementation of `Scripting.TextStream`.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/textstream-object
'
' @module TextStream
' @author Andrew Pullon | andrew.pullon@pkfh.co.nz | andrewcpullon@gmail.com
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder Scripting
'@ignoremodule ProcedureNotUsed
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

#If Not Mac And Not UseScriptingTextStreamIfAvailable Then

Private Declare PtrSafe Sub ByteCopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Dest As Any, Source As Any, ByVal Count As Long)

#End If

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingTextStreamIfAvailable = False

#If Mac Or Not UseScriptingTextStreamIfAvailable Then

Private Type TTextStream
    FileID As Long
    Mode As IOMode
    Format As Tristate
    FileLines() As Variant
    LineNo As Long
    ColumnNo As Long
End Type

#Else

Private Type TTextStream
    TextStream As Object
End Type

#End If

Private This As TTextStream

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

''
' Read-only property that returns True if the file pointer immediately precedes the end-of-line marker in a TextStream file; False if it does not.
'
' @property AtEndOfLine
' @type {Boolean}
''
Public Property Get AtEndOfLine() As Boolean
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    If StrPtr(This.FileLines(This.LineNo)) = 0 Then
        AtEndOfLine = True
    ElseIf This.ColumnNo > UBound(This.FileLines(This.LineNo)) Then
        AtEndOfLine = True
    Else
        AtEndOfLine = (This.FileLines(This.LineNo)(This.ColumnNo) = 10 Or This.FileLines(This.LineNo)(This.ColumnNo) = 13)
    End If
#Else
    AtEndOfLine = This.TextStream.AtEndOfLine
#End If
End Property

''
' Read-only property that returns True if the file pointer is at the end of a TextStream file; False if it is not.
'
' @property AtEndOfStream
' @type {Boolean}
''
Public Property Get AtEndOfStream() As Boolean
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    AtEndOfStream = (This.LineNo = UBound(This.FileLines) And AtEndOfLine)
#Else
    AtEndOfStream = This.TextStream.AtEndOfStream
#End If
End Property

''
' Read-only property that returns the column number of the current character position in a TextStream file.
'
' @property Column
' @type {Long}
''
Public Property Get Column() As Long
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If This.Format = TristateTrue Then
        Column = Application.WorksheetFunction.RoundUp(This.ColumnNo / 2, 0)
    Else
        Column = This.ColumnNo
    End If
#Else
    Column = This.TextStream.Column
#End If
End Property

''
' Read-only property that returns the current line number in a TextStream file.
'
' @property Line
' @type {Long}
''
Public Property Get Line() As Long
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    Line = This.LineNo
#Else
    Line = This.TextStream.Line
#End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' VBA implementation of TextStream `Close` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/close-method-textstream-object
'
' Closes an open TextStream file.
'
' @method CloseFile
''
Public Sub CloseFile()
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.FileID = -1 Then
        Close This.FileID
        This.FileID = -1
    End If
#Else
    If Not This.TextStream Is Nothing Then This.TextStream.Close
#End If
End Sub

''
' VBA-only method for TextStream, to provide for initialisation of file. Technically this should be
' using an interface to expose certain methods only (and this method should be private), but at this
' stage I don't want to add additional classes to achieve this.
'
' This method mimics `FileSystemObject.OpenTextFile`.
'
' @method OpenFile
' @param {String} FilePath
' @param {IOMode} Mode
' @param {Tristate} Format
''
Public Sub OpenFile(ByVal FilePath As String, ByVal Mode As IOMode, ByVal Format As Tristate)
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    This.FileID = VBA.FreeFile
    This.Mode = Mode
    This.Format = Format
    Select Case This.Mode
    Case IOMode.ForAppending
        Open FilePath For Binary Access Write As #This.FileID
        Seek This.FileID, VBA.LOF(This.FileID) + 1
    Case IOMode.ForWriting
        Open FilePath For Output As #This.FileID: Close #This.FileID ' Open then close file, to erase current content.
        Open FilePath For Binary Access Write As #This.FileID
        ' Write BOM if Unicode format.
        If This.Format = TristateTrue Then
            Dim txt_BOM(0 To 1) As Byte
            txt_BOM(0) = 255: txt_BOM(1) = 254
            Put #This.FileID, , txt_BOM
        End If
    Case IOMode.ForReading
        ' Read file contents to bytes.
        Dim txt_Bytes() As Byte
        Open FilePath For Binary Access Read As #This.FileID
        If Not VBA.LOF(This.FileID) = 0 Then ReDim txt_Bytes(1 To VBA.LOF(This.FileID))
        Get #This.FileID, 1, txt_Bytes
        ' Split bytes into lines, removing BOM if Unicode format.
        If This.Format = TristateTrue Then
            txtstr_BytesToLineBytes txtstr_RemoveBOM(txt_Bytes), This.FileLines
        Else
            txtstr_BytesToLineBytes txt_Bytes, This.FileLines
        End If
    End Select
#Else
    Dim txtstr_FSO As Object
    Set txtstr_FSO = CreateObject("Scripting.FileSystemObject")
    Set This.TextStream = txtstr_FSO.OpenTextFile(FilePath, Mode, True, Format)
#End If
End Sub

''
' VBA implementation of TextStream `Read` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/read-method
'
' Reads a specified number of characters from a TextStream file and returns the resulting string.
'
' @method Read
' @param {Long} Characters | Number of characters that you want to read from the file.
' @return {String}
''
Public Function Read(ByVal Characters As Long) As String
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    If Me.AtEndOfStream Then Err.Raise 62, , "Input past end of file"
    
    Dim txt_CharsToRead As Long
    Dim txt_CharsRead As Long
    Dim txt_Read() As Byte
    Dim txt_Long As Long
    
    txt_CharsToRead = Characters * VBA.IIf(This.Format = TristateTrue, 2, 1)
    
    ReDim txt_Read(1 To txt_CharsToRead)
    txt_CharsRead = 0
    Do Until txt_CharsRead = txt_CharsToRead
        ' Read characters from current line.
        For txt_Long = This.ColumnNo To UBound(This.FileLines(This.LineNo))
            txt_Read(txt_CharsRead + 1) = This.FileLines(This.LineNo)(txt_Long)
            txt_CharsRead = txt_CharsRead + 1
            If txt_CharsRead = txt_CharsToRead Then Exit For
        Next txt_Long
        ' Update column number, then move to next line if at the end.
        This.ColumnNo = txt_Long + 1
        If AtEndOfStream Then: This.ColumnNo = This.ColumnNo - 1: Exit Do
        If This.ColumnNo > UBound(This.FileLines(This.LineNo)) Then Me.SkipLine
    Loop
    Read = VBA.Replace(txtstr_ByteArrayToString(txt_Read), VBA.Chr$(0), vbNullString)
#Else
    Read = This.TextStream.Read(Characters)
#End If
End Function

''
' VBA implementation of TextStream `ReadAll` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/readall-method
'
' Reads an entire Text file and returns the resulting string.
'
' @method ReadFile
' @param {String} FilePath | Name of file to be read.
' @return {String}
''
Public Function ReadAll() As String
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    If Me.AtEndOfStream Then Err.Raise 62, , "Input past end of file"
    
    Dim txt_ReadAll() As Byte
    Dim txt_CharCount As Long
    Dim txt_LineNo As Long
    Dim txt_ColNo As Long
    txt_CharCount = 1
    
    For txt_LineNo = This.LineNo To UBound(This.FileLines)
        If Not StrPtr(This.FileLines(txt_LineNo)) = 0 Then
            ReDim Preserve txt_ReadAll(1 To txt_CharCount + UBound(This.FileLines(txt_LineNo)) - This.ColumnNo)
            For txt_ColNo = This.ColumnNo To UBound(This.FileLines(txt_LineNo))
                txt_ReadAll(txt_CharCount) = This.FileLines(txt_LineNo)(txt_ColNo)
                txt_CharCount = txt_CharCount + 1
            Next txt_ColNo
            This.ColumnNo = 1 ' Reset ColumnNo for next Line.
        End If
    Next txt_LineNo
    
    This.LineNo = txt_LineNo - 1
    If StrPtr(This.FileLines(This.LineNo)) = 0 Then ' Check if array is empty (file ends with blank line).
       This.ColumnNo = 1
    Else
        This.ColumnNo = UBound(This.FileLines(This.LineNo)) + 1
    End If
    
    ReadAll = txtstr_ByteArrayToString(txt_ReadAll)
#Else
    ReadAll = This.TextStream.ReadAll
#End If
End Function

''
' VBA implementation of TextStream `ReadLine` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/readline-method
'
' Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.
'
' @method ReadLine
' @return {String}
''
Public Function ReadLine() As String
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    If Me.AtEndOfStream Then Err.Raise 62, , "Input past end of file"
    
    Dim txt_ReadLine() As Byte
    Dim txt_ColNo As Long
    Dim txt_CharCount As Long
    
    ReDim txt_ReadLine(1 To UBound(This.FileLines(This.LineNo)) - (This.ColumnNo - 1))
    txt_CharCount = 1
    
    For txt_ColNo = This.ColumnNo To UBound(This.FileLines(This.LineNo))
        txt_ReadLine(txt_CharCount) = This.FileLines(This.LineNo)(txt_ColNo)
        txt_CharCount = txt_CharCount + 1
    Next txt_ColNo
    
    This.LineNo = This.LineNo + 1
    This.ColumnNo = 1
    
    ' Skip new line character when reading an entire line.
    ReadLine = VBA.Replace(VBA.Replace(VBA.Replace(txtstr_ByteArrayToString(txt_ReadLine), VBA.Chr$(13) & VBA.Chr$(10), vbNullString), VBA.Chr$(13), vbNullString), VBA.Chr$(10), vbNullString)
#Else
    ReadLine = This.TextStream.ReadLine
#End If
End Function

''
' VBA implementation of TextStream `Skip` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/skip-method
'
' Skips a specified number of characters when reading a TextStream file.
'
' @method Skip
' @param {Long} Characters | Number of characters to skip when reading a file.
''
Public Sub Skip(ByVal Characters As Long)
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    If Me.AtEndOfStream Then Err.Raise 62, , "Input past end of file"
    
    Dim txt_CharsToRead As Long
    Dim txt_CharsSkipped As Long
    Dim txt_Long As Long
    txt_CharsToRead = Characters * VBA.IIf(This.Format = TristateTrue, 2, 1)
    txt_CharsSkipped = 0
    
    
    Do Until txt_CharsSkipped = txt_CharsToRead
        For txt_Long = This.ColumnNo To UBound(This.FileLines(This.LineNo))
            txt_CharsSkipped = txt_CharsSkipped + 1
            If txt_CharsSkipped = txt_CharsToRead Then Exit For
        Next txt_Long
        This.ColumnNo = txt_Long + 1
        If AtEndOfStream Then: This.ColumnNo = This.ColumnNo - 1: Exit Do
        If This.ColumnNo > UBound(This.FileLines(This.LineNo)) Then Me.SkipLine
    Loop
#Else
    This.TextStream.Skip Characters
#End If
End Sub

''
' VBA implementation of TextStream `SkipLine` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/skipline-method
'
' Skips the next line when reading a TextStream file.
'
' @method SkipLine
''
Public Sub SkipLine()
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If Not This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    If Me.AtEndOfStream Then Err.Raise 62, , "Input past end of file"
    This.LineNo = This.LineNo + 1
    This.ColumnNo = 1
#Else
    This.TextStream.SkipLine
#End If
End Sub

''
' VBA implementation of TextStream `Write` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/write-method
'
' Writes a specified string to a Text file, overwriting existing content.
'
' @method WriteStr
' @param {String} Text | Text to be written to file.
''
Public Sub WriteStr(ByVal Text As String)
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    Put #This.FileID, , txtstr_StringToByteArray(Text)
#Else
    This.TextStream.Write Text
#End If
End Sub

''
' VBA implementation of TextStream `WriteBlankLines` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/writeblanklines-method
'
' @method WriteBlankLines
' @param {Long} Lines | Number of newline characters you want to write to the file.
''
Public Sub WriteBlankLines(ByVal Lines As Long)
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    
    Dim txt_Write As String
    Dim txt_Long As Long
    txt_Write = VBA.Space$(Lines * 2)
    For txt_Long = 1 To Lines * 2 Step 2
        Mid(txt_Write, txt_Long, 2) = VBA.Chr$(13) & VBA.Chr$(10)
    Next txt_Long
    
    Put #This.FileID, , txtstr_StringToByteArray(txt_Write)
#Else
    This.TextStream.WriteBlankLines Lines
#End If
End Sub

''
' VBA implementation of TextStream `WriteLine` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/writeline-method
'
' @method WriteLine
' @param {String} Text| Optional. The text you want to write to the file. If omitted, a newline character is written to the file.
''
Public Sub WriteLine(Optional ByVal Text As String)
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    If This.Mode = ForReading Then Err.Raise 54, , "Bad file mode"
    Put #This.FileID, , txtstr_StringToByteArray(Text & VBA.Chr$(13) & VBA.Chr$(10))
#Else
    This.TextStream.WriteLine Text
#End If
End Sub

' ============================================= '
' Private Methods
' ============================================= '

#If Mac Or Not UseScriptingTextStreamIfAvailable Then
''
' Removes BOM from the start of Unicode Byte Array.
'
' @method txtstr_RemoveBOM
' @param {Byte} Bytes | Byte array to remove BOM from.
' @return {Byte}
''
Private Function txtstr_RemoveBOM(ByRef Bytes() As Byte) As Byte()
    If Not StrPtr(Bytes) = 0 Then
        If UBound(Bytes) > 1 Then
            ' Check for BOM & remove if necessary.
            If Bytes(1) = 255 And Bytes(2) = 254 Then
                Dim txt_Return() As Byte
                Dim txt_Length As Long
                Dim txt_Long As Long
                txt_Length = UBound(Bytes) ' Get length of byte array.
                
                If txt_Length = 2 Then
                    ' If array only has the BOM, return empty byte array.
                    txtstr_RemoveBOM = txt_Return
                Else
                    ' Recreate source bytes without BOM.
                    ReDim txt_Return(1 To txt_Length - 2)
                    For txt_Long = 3 To txt_Length
                        txt_Return(txt_Long - 2) = Bytes(txt_Long)
                    Next txt_Long
                    ' Return modified bytes.
                    txtstr_RemoveBOM = txt_Return
                End If
            Else
                ' Return original bytes
                txtstr_RemoveBOM = Bytes
            End If
        Else
            ' Return original bytes
            txtstr_RemoveBOM = Bytes
        End If
    Else
        ' Return original bytes
        txtstr_RemoveBOM = Bytes
    End If
End Function

''
' Split file bytes into lines, looking for vbNewLine/vbCrLf (Chr(13) + Chr(10)).
'
' @method txtstr_BytesToLineBytes
' @param {Byte} | All bytes in opened file.
' @param[out] {Array} | Array to save each line of bytes to.
''
Private Sub txtstr_BytesToLineBytes(ByRef FileBytes() As Byte, ByRef LineBytes() As Variant)
    Dim txt_LineBytes() As Byte
    Dim txt_ByteNo As Long
    Dim txt_LineNo As Long
    Dim txt_LineBreakStampNo As Long
    
    txt_LineNo = 1
    txt_LineBreakStampNo = 0
    
    If Not StrPtr(FileBytes) = 0 Then
        ReDim txt_LineBytes(1 To 1)
        txt_LineBytes(1) = FileBytes(1)
        
        For txt_ByteNo = 2 To UBound(FileBytes)
            ReDim Preserve txt_LineBytes(1 To txt_ByteNo - txt_LineBreakStampNo)
            txt_LineBytes(txt_ByteNo - txt_LineBreakStampNo) = FileBytes(txt_ByteNo)
            Select Case This.Format
            Case Tristate.TristateFalse ' ASCII
                If FileBytes(txt_ByteNo - 1) = 13 And FileBytes(txt_ByteNo) = 10 Then
                    txt_LineBreakStampNo = txt_ByteNo
                    bytes_SaveLine txt_LineBytes, LineBytes, txt_LineNo
                End If
            Case Tristate.TristateTrue ' Unicode.
                If txt_ByteNo > 3 Then
                    If FileBytes(txt_ByteNo - 3) = 13 And FileBytes(txt_ByteNo - 2) = 0 And FileBytes(txt_ByteNo - 1) = 10 And FileBytes(txt_ByteNo) = 0 Then
                        txt_LineBreakStampNo = txt_ByteNo
                        bytes_SaveLine txt_LineBytes, LineBytes, txt_LineNo
                    End If
                End If
            End Select
        Next txt_ByteNo
        bytes_SaveLine txt_LineBytes, LineBytes, txt_LineNo
    Else
        ' Return one line with empty byte array.
        ReDim LineBytes(1 To 1)
        LineBytes(1) = FileBytes
        Exit Sub
    End If
End Sub
Private Sub bytes_SaveLine(ByRef LineBytes() As Byte, ByRef LineArray() As Variant, ByRef LineNo As Long)
    ReDim Preserve LineArray(1 To LineNo)
    LineArray(LineNo) = LineBytes
    Erase LineBytes
    LineNo = LineNo + 1
End Sub

''
' Converts byte array to string. Correctly handles Unicode bytes.
'
' @method txt_ByteArrayToString
' @param {Byte} Bytes | Byte array to convert to string.
' @return {String}
''
Private Function txtstr_ByteArrayToString(ByRef Bytes() As Byte) As String
    ' Variables.
    Dim txt_Length As Long

    If StrPtr(Bytes) = 0 Then
        txt_Length = 0
    Else
        txt_Length = UBound(Bytes)
    End If

    Select Case This.Format
    Case Tristate.TristateFalse ' ASCII
        txtstr_ByteArrayToString = VBA.StrConv(Bytes, vbUnicode)
    Case Tristate.TristateTrue ' Unicode
        ' Account for terminating null
        If (txt_Length Mod 2) Then txt_Length = txt_Length - 1
        ' Set up a buffer to recieve the string.
        txtstr_ByteArrayToString = VBA.String$(txt_Length / 2, 0)
        ' Copy to string.
        ByteCopyMemory ByVal StrPtr(txtstr_ByteArrayToString), Bytes(1), txt_Length
    End Select
End Function

''
' Converts string to byte array. Correctly handles Unicode.
'
' @method dpapi_StringToByteArray
' @param {String} Text | String to convert to byte array.
' @return {Byte}
''
Private Function txtstr_StringToByteArray(ByVal Text As String) As Byte()
    ' Variables.
    Dim txt_Bytes() As Byte
    Dim txt_Length As Long
    
    Select Case This.Format
    Case Tristate.TristateFalse ' ASCII
        txt_Bytes = VBA.StrConv(Text, vbFromUnicode)
    Case Tristate.TristateTrue ' Unicode
        ' Number of bytes.
        txt_Length = VBA.LenB(Text)
        If txt_Length = 0 Then Exit Function
        ReDim txt_Bytes(txt_Length - 1)
        ' Copy characters from string to byte array.
        ByteCopyMemory txt_Bytes(0), ByVal StrPtr(Text), txt_Length
    End Select
    
    txtstr_StringToByteArray = txt_Bytes
End Function

#End If

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

Private Sub Class_Terminate()
    Me.CloseFile
End Sub

Private Sub Class_Initialize()
#If Mac Or Not UseScriptingTextStreamIfAvailable Then
    This.FileID = -1
    This.LineNo = 1
    This.ColumnNo = 1
#End If
End Sub
