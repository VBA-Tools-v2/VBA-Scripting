VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' FileSystemObject | STATUS: In Development.
' (c) RadiusCore Ltd - https://radiuscore.co.nz/
'
' Provides access to a computer's file system.
'
' Native VBA implementation of `Scripting.FileSystemObject`
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'
' @module FileSystemObject
' @author Andrew Pullon | andrew.pullon@pkfh.co.nz | andrewcpullon@gmail.com
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder Scripting
'@ignoremodule ProcedureNotUsed
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Public Enum IOMode
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum

Public Enum Tristate
    TristateUseDefault = -2
    TristateMixed = -2
    TristateFalse = 0
    TristateTrue = -1
End Enum

'Public Enum FileAttribute
'    Normal = 0
'    ReadOnly = 1
'    Hidden = 2
'    System = 4
'    Volume = 8
'    Directory = 16
'    Archive = 32
'    Alias = 1024
'    Compressed = 2048
'End Enum

' ============================================= '
' Public Methods
' ============================================= '

''
' VBA implementation of FileSystemObject `BuildPath` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/buildpath-method
'
' Combines a folder path and the name of a folder or file and returns the combination with valid path separators.
'
' @method BuildPath
' @param {String} Path | Existing path with which `Name` is combined. Path can be absolute or relative and need not specify an existing folder.
' @param {String} Name | Name of a folder or file being appended to the existing `Path`.
' @return {String}
''
Public Function BuildPath(ByVal Path As String, ByVal Name As String) As String
    ' Variables.
    Dim fso_Path As String
    Dim fso_Name As String
    Dim fso_Separator As String
    fso_Path = Path
    fso_Name = Name
    
    ' Determine correct path separator.
    ' Check path for valid sepatator first. If one is not found, check the Name.
    If VBA.InStr(fso_Path, "/") Then
        fso_Separator = "/"
    ElseIf VBA.InStr(fso_Path, "\") Then
        fso_Separator = "\"
    ElseIf VBA.InStr(fso_Name, "/") Then
        fso_Separator = "/"
    ElseIf VBA.InStr(fso_Name, "\") Then
        fso_Separator = "/"
#If Mac Then
        ' Mac, use default Mac path separator.
        fso_Separator = ":"
#Else
        ' Windows, use default Windows path separator.
        fso_Separator = "\"
#End If
    End If
    
    ' Add path separator to path if required.
    If Not fso_Path = vbNullString And Not fso_Name = vbNullString Then
        If Not (VBA.Right$(fso_Path, 1) = fso_Separator) Then
            fso_Path = fso_Path & fso_Separator
        End If
    End If
    
    ' Remove path separator from name if required. (check for ALL separators).
    If VBA.Left$(fso_Name, 1) = "/" Or VBA.Left$(fso_Name, 1) = "\" Then
        fso_Name = VBA.Right$(fso_Name, VBA.Len(fso_Name) - 1)
    End If
#If Mac Then
    If VBA.Left$(fso_Name, 1) = ":" Then
        fso_Name = VBA.Right$(fso_Name, VBA.Len(fso_Name) - 1)
    End If
#End If
    
    ' Create path
    BuildPath = fso_Path & fso_Name
End Function

''
' VBA implementation of FileSystemObject `CopyFile` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copyfile-method
'
' Copies one file from one location to another.
'
' TODO - support wild-card in `Source` parameter to enable copying multiple files.
'
' @method BuildPath
' @param {String} Source | Character string file specification, for the file to be copied.
' @param {String} Destination | Character string destination where the file or files from source are to be copied.
' @param {Boolean} Overwrite | Optional. Boolean value that indicates if existing files are to be overwritten.
'                              If True, files are overwritten; if False, they are not. The default is True.
'                              Note that CopyFile will fail if destination has the read-only attribute set, regardless of the value of overwrite.
''
Public Sub CopyFile(ByVal Source As String, ByVal Destination As String, Optional ByVal Overwrite As Boolean = True)
    Dim fso_Source As String
    Dim fso_Destination As String
    
    fso_Source = Source
    fso_Destination = Destination
    
    If Not VBA.InStr(fso_Destination, ".") > 0 Then
        ' Extract file name from source if a new file name has not been specified.
        If VBA.InStrRev(fso_Source, "/") > 0 Then
            fso_Source = VBA.Right$(fso_Source, VBA.Len(fso_Source) - VBA.InStrRev(fso_Source, "/"))
        ElseIf VBA.InStrRev(fso_Source, "\") > 0 Then
           fso_Source = VBA.Right$(fso_Source, VBA.Len(fso_Source) - VBA.InStrRev(fso_Source, "\"))
        Else
            Err.Raise 75, "CopyFile", "Path/file access error"
        End If
        ' Update destination to include original file name.
        fso_Destination = BuildPath(fso_Destination, fso_Source)
    End If
    
    If Not Overwrite Then
        If FileExists(fso_Destination) Then Err.Raise 58, "CopyFile", "File already exists."
    End If
    
    ' Copy file.
    VBA.FileCopy Source, fso_Destination
End Sub

'TODO - CopyFolder.

''
' VBA implementation of FileSystemObject `CreateFolder` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createfolder-method
'
' Creates a folder.
'
' NOTE: upgraded to create folders recursively (which the original method does not do).
'
' @method CreateFolder
' @param {String} FolderName | String expression that identifies the folder to create.
''
Public Sub CreateFolder(ByVal FolderName As String)
    If FolderName = vbNullString Then Err.Raise 5, , "Invalid procedure call or argument"
    If Me.FolderExists(FolderName) Then Err.Raise 58, , "File already exists."
    
    
    ' Recursively create the parent folder (if it doesn't already exist).
    If Not Me.FolderExists(Me.GetParentFolderName(FolderName)) Then Me.CreateFolder Me.GetParentFolderName(FolderName)
    
    ' Create folder target folder.
    VBA.MkDir FolderName
End Sub

''
' Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.
'
' @method CreateTextFile
' @param {String} FileName | String expression that identifies the file to create.
' @param {Boolean} Overwrite | Indicates if an existing file can be overwritten. The value is True if the file can be overwritten; False if it can't be overwritten. If omitted, existing files can be overwritten.
' @param {Boolean} Unicode |  Indicates whether the file is created as a Unicode or ASCII file. The value is True if the file is created as a Unicode file; False if it's created as an ASCII file. If omitted, an ASCII file is assumed.
' @return {TextStream}
''
Public Function CreateTextFile(ByVal FileName As String, Optional ByVal Overwrite As Boolean = True, Optional ByVal Unicode As Boolean = False) As TextStream
    If Not Overwrite Then
        If Me.FileExists(FileName) Then Err.Raise 58, , "File already exists"
    Else
        If Me.FileExists(FileName) Then Me.DeleteFile FileName
    End If
    Dim fso_TextStream As TextStream
    Set fso_TextStream = New TextStream
    fso_TextStream.OpenFile FileName, ForWriting, VBA.IIf(Unicode, TristateTrue, TristateFalse)
    Set CreateTextFile = fso_TextStream
End Function

''
' VBA implementation of FileSystemObject `DeleteFile` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletefile-method
'
' Deletes a specified file.
'
' @method DeleteFile
' @param {String} FileSpec | The name of the file to delete. The filespec can contain wildcard characters in the last path component.
' @param {Boolean} Force | Optional. Boolean value that is True if folders with the read-only attribute set are to be deleted; False (default) if they are not.
''
Public Sub DeleteFile(ByVal FileSpec As String, Optional ByVal Force As Boolean = False)
    ' Skip if Force = False and file is read-only.
    If Not Force And (VBA.GetAttr(FileSpec) And vbReadOnly) > 0 Then
        Exit Sub
    End If
    
    VBA.Kill FileSpec
End Sub

''
' VBA implementation of FileSystemObject `DeleteFolder` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletefolder-method
'
' Deletes a specified folder and its contents.
'
' @method DeleteFolder
' @param {String} FolderSpec | The name of the folder to delete. The folderspec can contain wildcard characters in the last path component.
' @param {Boolean} Force | Optional. Boolean value that is True if folders with the read-only attribute set are to be deleted; False (default) if they are not.
''
Public Sub DeleteFolder(ByVal FolderSpec As String, Optional ByVal Force As Boolean = False)
    On Error GoTo fso_ErrorHandling
    
    Dim fso_Folder As String
    
    If Not FolderExists(FolderSpec) Then Exit Sub
    
    ' Skip if Force = False and folder is read-only.
    If Not Force And (VBA.GetAttr(FolderSpec) And vbReadOnly) > 0 Then
        'Debug.Print "Skipping Read-Only folder: " & FolderSpec
        Exit Sub
    End If
    
    ' Delete all files.
    On Error Resume Next ' Ignore error if no files in folder.
        'Debug.Print "Deleting files: " & BuildPath(FolderSpec, "*.*")
        VBA.Kill BuildPath(FolderSpec, "*.*")
    On Error GoTo fso_ErrorHandling
    
    ' Recursively loop through all folders and delete contents.
    fso_Folder = VBA.Dir(BuildPath(FolderSpec, vbNullString), vbDirectory)
    Do While Not fso_Folder = vbNullString
        On Error GoTo fso_FixDir
            fso_Folder = VBA.Dir()
        On Error GoTo fso_ErrorHandling
        If Not (fso_Folder = "." Or fso_Folder = ".." Or fso_Folder = vbNullString) Then
            DeleteFolder BuildPath(FolderSpec, fso_Folder), Force
        End If
    Loop
    
    'Debug.Print "Deleting folder: " & FolderSpec
    VBA.RmDir FolderSpec
    
    Exit Sub
    
fso_FixDir:
    fso_Folder = VBA.Dir(BuildPath(FolderSpec, vbNullString), vbDirectory)
    Resume Next
fso_ErrorHandling:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

''
' VBA implementation of FileSystemObject `FileExists` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fileexists-method
'
' Returns `True` if a specified file exists; `False` if it does not.
'
' @method FileExists
' @param {String} FileSpec | The name of the file whose existence is to be determined.
' @return {Boolean}
''
Public Function FileExists(ByVal FileSpec As String) As Boolean
    Dim fso_File As String
    fso_File = FileSpec
    FileExists = False
    
    If VBA.Right$(fso_File, 1) = "/" Or VBA.Right$(fso_File, 1) = "\" Or VBA.Right$(fso_File, 1) = ":" Then
        fso_File = VBA.Left$(fso_File, VBA.Len(fso_File) - 1)
    End If
    
    If VBA.Len(VBA.Dir(fso_File)) > 0 Then
        FileExists = True
    End If
End Function

''
' VBA implementation of FileSystemObject `FolderExists` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folderexists-method
'
' Returns `True` if a specified folder exists; `False` if it does not.
'
' @method FolderExists
' @param {String} FolderSpec | The name of the folder whose existence is to be determined.
' @return {Boolean}
''
Public Function FolderExists(ByVal FolderSpec As String) As Boolean
    Dim fso_Folder As String
    fso_Folder = FolderSpec
    FolderExists = False
    
    If Not (VBA.Right$(fso_Folder, 1) = "/") And Not (VBA.Right$(fso_Folder, 1) = "\") And Not (VBA.Right$(fso_Folder, 1) = ":") Then
        If VBA.InStr(fso_Folder, "/") > 0 Then
            fso_Folder = fso_Folder & "/"
        ElseIf VBA.InStr(fso_Folder, "\") > 0 Then
            fso_Folder = fso_Folder & "\"
        ElseIf VBA.InStr(fso_Folder, ":") > 0 Then
            fso_Folder = fso_Folder & ":"
        Else
            FolderExists = False
            Exit Function
        End If
    End If
    
    If VBA.Len(VBA.Dir(fso_Folder, vbDirectory)) > 0 Then
        FolderExists = True
    End If
End Function

''
' VBA implementation of FileSystemObject `GetBaseName` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getbasename-method
'
' Returns a string containing the base name of the last component, less any file extension, in a path.
'
' @method GetBaseName
' @param {String} Path | The path specification for the component whose base name is to be returned.
' @return {String}
''
Public Function GetBaseName(ByVal Path As String) As String
    ' TEMP
    Dim fso_Object As Object
    Set fso_Object = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso_Object.GetBaseName(Path)
End Function

''
' VBA implementation of FileSystemObject `GetExtensionName` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getextensionname-method
'
' Returns a string containing the extension name for the last component in a path.
'
' @method GetExtensionName
' @param {String} Path | The path specification for the component whose extension name is to be returned.
' @return {String}
''
Public Function GetExtensionName(ByVal Path As String) As String
    ' TODO - need to check for '.' only in the LAST part of the path.
    If Not VBA.InStr(Path, ".") > 0 Then Exit Function
    GetExtensionName = VBA.Right$(Path, VBA.Len(Path) - VBA.InStrRev(Path, "."))
End Function

''
' VBA implementation of FileSystemObject `GetFile` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfile-method
'
' Returns a File object corresponding to the file in a specified path.
'
' @method GetFileName
' @param {String} FileSpec | The path (absolute or relative) to a specific file.
' @return {File}
''
Public Function GetFile(ByVal FileSpec As String) As Object
    ' TEMP
    Dim fso_Object As Object
    Set fso_Object = CreateObject("Scripting.FileSystemObject")
    Set GetFile = fso_Object.GetFile(FileSpec)
End Function

''
' VBA implementation of FileSystemObject `GetFileName` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfilename-method-visual-basic-for-applications
'
' Returns the last component of a specified path that is not part of the drive specification.
'
' @method GetFileName
' @param {String} PathSpec | The path (absolute or relative) to a specific file.
' @return {String}
''
Public Function GetFileName(ByVal PathSpec As String) As String
    ' TEMP
    Dim fso_Object As Object
    Set fso_Object = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso_Object.GetFileName(PathSpec)
End Function

''
' VBA implementation of FileSystemObject `GetFolder` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
'
' Returns a Folder object corresponding to the folder in a specified path.
'
' @method GetFolder
' @param {String} FolderSpec | The folderspec is the path (absolute or relative) to a specific folder.
' @return {String}
''
Public Function GetFolder(ByVal FolderSpec As String) As Object
    ' TEMP
    If Not Me.FolderExists(FolderSpec) Then
        Debug.Print "Raise Error Here"
    End If
    
    Dim fso_Object As Object
    Set fso_Object = CreateObject("Scripting.FileSystemObject")
    Set GetFolder = fso_Object.GetFolder(FolderSpec)
End Function

''
' VBA implementation of FileSystemObject `GetParentFolderName` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getparentfoldername-method
'
' Returns a string containing the name of the parent folder of the last component in a specified path.
'
' @method FolderExists
' @param {String} Path | The path specification for the component whose parent folder name is to be returned.
' @return {String}
''
Public Function GetParentFolderName(ByVal Path As String) As String
    If VBA.InStr(Path, "/") > 0 Then
        GetParentFolderName = VBA.Mid$(Path, 1, VBA.InStrRev(Path, "/") - 1)
    ElseIf VBA.InStr(Path, "\") > 0 Then
        GetParentFolderName = VBA.Mid$(Path, 1, VBA.InStrRev(Path, "\") - 1)
    End If
End Function

''
' VBA implementation of FileSystemObject `MoveFile` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/movefile-method
'
' Moves one or more files from one location to another.
'
' @method MoveFile
' @param {String} Source | The path to the file or files to be moved. The source argument string can contain wildcard characters in the last path component only.
' @param {String} Destination | The path where the file or files are to be moved. The destination argument can't contain wildcard characters.
''
Public Sub MoveFile(ByVal Source As String, ByVal Destination As String)
    ' TEMP
    Dim fso_Object As Object
    Set fso_Object = CreateObject("Scripting.FileSystemObject")
    fso_Object.MoveFile Source, Destination
End Sub

''
' VBA implementation of FileSystemObject `OpenTextFile` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
'
' Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.
'
' @method OpenTextFile
' @param {String} FileName | String expression that identifies the file to open.
' @param {IOMode} Mode | Optional. Indicates input/output mode. Can be one of three constants: ForReading, ForWriting, or ForAppending.
' @param {Boolean} Create | Optional. Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created; False if it isn't created. The default is False.
' @param {Tristate} Format | Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.
' @return {TextStream}
''
Public Function OpenTextFile(ByVal FileName As String, Optional ByVal Mode As IOMode, Optional ByVal Create As Boolean = False, Optional ByVal Format As Tristate = Tristate.TristateFalse) As TextStream
    If Not Create Then
        If Not Me.FileExists(FileName) Then Err.Raise 53, , "File not found"
    End If
    Dim fso_TextStream As TextStream
    Set fso_TextStream = New TextStream
    fso_TextStream.OpenFile FileName, Mode, Format
    Set OpenTextFile = fso_TextStream
End Function
