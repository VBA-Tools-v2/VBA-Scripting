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
' Notes on speed:
' - Checking if string is empty is ~2 times faster using `MyString = vbNullString` vs `VBA.Len(MyString) = 0`.
' - Checking the first character of a string is ~1.5 times faster using `VBA.Asc(MyString) = {AscCode}` vs `VBA.Left$(MyString,1) = {Letter}`.
'    - NOTE: Asc method needs to have vbNullString check first, as it errors with an empty string. It is still faster even with this check.
' - If `VBA.Mid$` can be used to replace a `VBA.Right$` call (which also requires a `VBA.Len` call) it will be ~1.5 times faster.
' - Leaving params as implicitly `ByRef` is ~2 times faster then specifying `ByVal`. Although methods need to be careful not to modify the provided
'   params or the implicit `ByRef` can cause bugs.
' - If checking for the existance of one string within another, leaving the If statement without a specific right-hand value is ~1.1 times faster.
'    - i.e. `If VBA.InStr(MyString, {Char}) Then` provides same effect (and is faster than) `If VBA.InStr(MyString, {Char}) > 0 Then`.
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

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingFileSystemObjectIfAvailable = False

#If Not Mac And UseScriptingFileSystemObjectIfAvailable Then

Private Type TFileSystemObject
    FileSystemObject As Object ' Scripting.FileSystemObject
End Type

Private This As TFileSystemObject

#End If

' --------------------------------------------- '
' Public Enums
' --------------------------------------------- '

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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    Dim fso_Separator As String

    ' Check path for valid sepatator first. If one is not found, check the Name.
    If VBA.InStr(Path, "/") Then
        fso_Separator = "/"
    ElseIf VBA.InStr(Path, "\") Then
        fso_Separator = "\"
    ElseIf VBA.InStr(Name, "/") Then
        fso_Separator = "/"
    ElseIf VBA.InStr(Name, "\") Then
        fso_Separator = "\"
    Else
        fso_Separator = Application.PathSeparator ' Use default file separator.
    End If

    ' Start result with `Path`.
    BuildPath = Path

    ' Add file separator (if not already present).
    If Not Path = vbNullString And Not Name = vbNullString Then
        If Not VBA.Right$(Path, 1) = fso_Separator Then
            BuildPath = BuildPath & fso_Separator
        End If
    End If

    ' Append `Name`, removing separator if present (check for ALL separators).
    If Not Name = vbNullString Then
        Select Case VBA.Asc(Name)
        Case 47, 92 ' ASCII Codes: `/`=47 | `\`=92
            BuildPath = BuildPath & VBA.Mid$(Name, 2)
        Case Else
            BuildPath = BuildPath & Name
        End Select
    End If
#Else
    BuildPath = This.FileSystemObject.BuildPath(Path, Name)
#End If
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
' @param {Boolean} Overwrite | Optional. Boolean value that indicates if existing files are to be overwritten. If True, files are overwritten; if False, they are not. The default is True.
'                              Note that CopyFile will fail if destination has the read-only attribute set, regardless of the value of overwrite.
''
Public Sub CopyFile(ByVal Source As String, ByVal Destination As String, Optional ByVal Overwrite As Boolean = True)
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    Dim fso_Destination As String
    
    ' If `Source` contains wildcard characters, or `Destination` ends with a path separator, it is assumed that
    ' `Destination` is an existing folder in which to copy matching files.
    If VBA.InStr(Me.GetFileName(Source), "*") Or (VBA.Right$(Destination, 1) = "/" Or VBA.Right$(Destination, 1) = "\") Then
        If Not Me.FolderExists(Destination) Then Err.Raise 76 ' Path not found.
        If VBA.InStr(Me.GetFileName(Source), "*") Then
            ' TODO - Recursively copy with wildcard.
        Else
            fso_Destination = Me.BuildPath(Destination, Me.GetFileName(Source)) ' Add file name to Destination.
        End If
    Else
        fso_Destination = Destination
    End If
    
    ' Non-wild card source, check if source file exists.
    If Not Me.FileExists(Source) Then Err.Raise 53 ' File not found.
    
    If Me.FileExists(fso_Destination) Then
        If Not Overwrite Then Err.Raise 58 ' File already exists.
        Me.DeleteFile fso_Destination, True ' Delete existing file.
    End If
    
    VBA.FileCopy Source, fso_Destination ' Copy file.
#Else
    This.FileSystemObject.CopyFile Source, Destination, Overwrite
#End If
End Sub

''
' VBA implementation of FileSystemObject `CopyFolder` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copyfolder-method
'
' Recursively copies a folder from one location to another.
'
' @method CopyFolder
' @param {String} Source | Character string folder specification, which can include wildcard characters, for one or more folders to be copied.
' @param {String} Destination | Character string destination where the folder and subfolders from source are to be copied. Wildcard characters are not allowed.
' @param {Boolean} Overwrite | Optional. Boolean value that indicates if existing folders are to be overwritten. If True, files are overwritten; if False, they are not. The default is True.
''
Public Sub CopyFolder(ByVal Source As String, ByVal Destination As String, Optional ByVal Overwrite As Boolean = True)
    ' TEMP
    Dim fso_Object As Object
    Set fso_Object = CreateObject("Scripting.FileSystemObject")
    fso_Object.CopyFolder Source, Destination, Overwrite
End Sub

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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If FolderName = vbNullString Then Err.Raise 5, , "Invalid procedure call or argument"
    If Me.FolderExists(FolderName) Then Err.Raise 58, , "File already exists."
    
    ' Recursively create the parent folder (if it doesn't already exist).
    If Not Me.FolderExists(Me.GetParentFolderName(FolderName)) Then Me.CreateFolder Me.GetParentFolderName(FolderName)
    
    ' Create folder target folder.
    VBA.MkDir FolderName
#Else
    ' Recursively create the parent folder (if it doesn't already exist).
    If Not Me.FolderExists(Me.GetParentFolderName(FolderName)) Then Me.CreateFolder Me.GetParentFolderName(FolderName)
    
    This.FileSystemObject.CreateFolder FolderName
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If Not Overwrite Then
        If Me.FileExists(FileName) Then Err.Raise 58, , "File already exists"
    Else
        If Me.FileExists(FileName) Then Me.DeleteFile FileName
    End If
    Dim fso_TextStream As TextStream
    Set fso_TextStream = New TextStream
    fso_TextStream.Initialise FileName, ForWriting, VBA.IIf(Unicode, TristateTrue, TristateFalse)
    Set CreateTextFile = fso_TextStream
#Else
    Set CreateTextFile = This.FileSystemObject.CreateTextFile(FileName, Overwrite, Unicode)
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If Not Force And (VBA.GetAttr(FileSpec) And vbReadOnly) > 0 Then Exit Sub           ' Skip if Force = False and file is read-only.
    If (VBA.GetAttr(FileSpec) And vbReadOnly) > 0 Then VBA.SetAttr FileSpec, vbNormal   ' Remove read-only if necessary.
    VBA.Kill FileSpec                                                                   ' Delete file.
#Else
    This.FileSystemObject.DeleteFile FileSpec, Force
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    On Error GoTo fso_ErrorHandling
    
    Dim fso_Folder As String
    
    If Not Me.FolderExists(FolderSpec) Then Exit Sub
    
    ' Skip if Force = False and folder is read-only.
    If Not Force And (VBA.GetAttr(FolderSpec) And vbReadOnly) > 0 Then
        Exit Sub
    End If
    
    ' Delete all files.
    On Error Resume Next ' Ignore error if no files in folder.
        VBA.Kill Me.BuildPath(FolderSpec, "*.*")
    On Error GoTo fso_ErrorHandling
    
    ' Recursively loop through all folders and delete contents.
    fso_Folder = VBA.Dir(Me.BuildPath(FolderSpec, Application.PathSeparator), vbDirectory)
    Do While Not fso_Folder = vbNullString
        On Error GoTo fso_FixDir
            fso_Folder = VBA.Dir()
        On Error GoTo fso_ErrorHandling
        If Not (fso_Folder = "." Or fso_Folder = ".." Or fso_Folder = vbNullString) Then
            DeleteFolder Me.BuildPath(FolderSpec, fso_Folder), Force
        End If
    Loop
    
    VBA.RmDir FolderSpec
    
    Exit Sub
    
fso_FixDir:
    fso_Folder = VBA.Dir(Me.BuildPath(FolderSpec, Application.PathSeparator), vbDirectory)
    Resume Next
fso_ErrorHandling:
    Err.Raise Err.Number, Err.Source, Err.Description
#Else
    This.FileSystemObject.DeleteFolder FolderSpec, Force
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If FileSpec = vbNullString Or VBA.Right$(FileSpec, 1) = "/" Or VBA.Right$(FileSpec, 1) = "\" Then Exit Function ' Return False if passed `FileSpec` that is blank or ending in file separator.
    FileExists = Not VBA.Dir(FileSpec) = vbNullString
#Else
    FileExists = This.FileSystemObject.FileExists(FileSpec)
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If FolderSpec = vbNullString Or Me.FileExists(FolderSpec) Then Exit Function ' Return False if passed `FolderSpec` that is blank or is a file.
    FolderExists = Not (VBA.Dir(FolderSpec, vbDirectory) = vbNullString)
#Else
    FolderExists = This.FileSystemObject.FolderExists(FolderSpec)
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    ' Get last component in path.
    GetBaseName = Me.GetFileName(Path)
    
    ' Remove file extension if present.
    If VBA.InStr(GetBaseName, ".") <> 0 Then
        GetBaseName = VBA.Mid$(GetBaseName, 1, VBA.InStrRev(GetBaseName, ".") - 1)
    End If
#Else
    GetBaseName = This.FileSystemObject.GetBaseName(Path)
#End If
End Function

''
' VBA implementation of FileSystemObject `GetDriveName` method.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getdrivename-method
'
' Returns a string containing the name of the drive for a specified path.
'
' @method GetDriveName
' @param {String} Path | The path specification for the component whose drive name is to be returned.
' @return {String}
''
Public Function GetDriveName(ByVal Path As String) As String
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If VBA.Len(Path) >= 2 Then
        ' ASCII Codes: `A`=65 | `Z`=90
        If VBA.Asc(Path) >= 65 And VBA.Asc(Path) <= 90 And VBA.Mid$(Path, 2, 1) = ":" Then
            GetDriveName = VBA.Left$(Path, 2)
        End If
    End If
#Else
    GetDriveName = This.FileSystemObject.GetDriveName(Path)
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    ' Get last component in path.
    GetExtensionName = Me.GetFileName(Path)

    ' Extract file extension if present.
    If VBA.InStr(GetExtensionName, ".") Then
        GetExtensionName = VBA.Mid$(GetExtensionName, VBA.InStrRev(GetExtensionName, ".") + 1)
    Else
        GetExtensionName = vbNullString
    End If
#Else
    GetExtensionName = This.FileSystemObject.GetExtensionName(Path)
#End If
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
    ' TEMP - Need to build `File` class for native implementation.
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    ' Variables.
    Dim fso_Backslash As Long
    Dim fso_Forwardslash As Long
    
    ' Trim file separator from path.
    Select Case VBA.Right$(PathSpec, 1)
    Case "\", "/"
        GetFileName = VBA.Mid$(PathSpec, 1, VBA.Len(PathSpec) - 1)
    Case Else
        GetFileName = PathSpec
    End Select
    
    ' Check if remaining path is a drive spec, exit if true.
    If VBA.Len(GetFileName) = 2 Then
        ' ASCII Codes: `A`=65 | `Z`=90
        If VBA.Asc(GetFileName) >= 65 And VBA.Asc(GetFileName) <= 90 And VBA.Right$(GetFileName, 1) = ":" Then
            GetFileName = vbNullString
            Exit Function
        End If
    End If
    
    ' Find position of file separators.
    fso_Backslash = VBA.InStrRev(GetFileName, "\")
    fso_Forwardslash = VBA.InStrRev(GetFileName, "/")
    
    ' Remove path to left of file separator.
    If fso_Backslash = 0 And fso_Forwardslash = 0 Then ' No file separator.
        Exit Function
    ElseIf fso_Backslash = 0 Then  ' Only Forwardslash separator.
        GetFileName = VBA.Mid$(GetFileName, fso_Forwardslash + 1)
    ElseIf fso_Forwardslash = 0 Then  ' Only Backwardslash separator.
        GetFileName = VBA.Mid$(GetFileName, fso_Backslash + 1)
    Else ' Both separators, use last one in path.
        GetFileName = VBA.Mid$(GetFileName, VBA.IIf(fso_Backslash > fso_Forwardslash, fso_Backslash, fso_Forwardslash) + 1)
    End If
#Else
    PathSpec = This.FileSystemObject.GetFileName(PathSpec)
#End If
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
    ' TEMP - Need to build `Folder` class for native implementation.
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
' @method GetParentFolderName
' @param {String} Path | The path specification for the component whose parent folder name is to be returned.
' @return {String}
''
Public Function GetParentFolderName(ByVal Path As String) As String
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If VBA.InStr(Path, "/") Then
        GetParentFolderName = VBA.Mid$(Path, 1, VBA.InStrRev(Path, "/") - 1)
    ElseIf VBA.InStr(Path, "\") Then
        GetParentFolderName = VBA.Mid$(Path, 1, VBA.InStrRev(Path, "\") - 1)
    End If
#Else
    GetParentFolderName = This.FileSystemObject.GetParentFolderName(Path)
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    ' TODO - Validate `Source` and `Destination` parameters.
    Name Source As Destination
#Else
    This.FileSystemObject.MoveFile Source, Destination
#End If
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
#If Mac Or Not UseScriptingFileSystemObjectIfAvailable Then
    If Not Create Then
        If Not Me.FileExists(FileName) Then Err.Raise 53, , "File not found"
    End If
    Dim fso_TextStream As TextStream
    Set fso_TextStream = New TextStream
    fso_TextStream.Initialise FileName, Mode, Format
    Set OpenTextFile = fso_TextStream
#Else
    Set OpenTextFile = This.FileSystemObject.OpenTextFile(FileName, Mode, Create, Format)
#End If
End Function

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

#If Not Mac And UseScriptingFileSystemObjectIfAvailable Then

Private Sub Class_Terminate()
    Set This.FileSystemObject = Nothing
End Sub

Private Sub Class_Initialize()
    Set This.FileSystemObject = CreateObject("Scripting.FileSystemObject")
End Sub

#End If
