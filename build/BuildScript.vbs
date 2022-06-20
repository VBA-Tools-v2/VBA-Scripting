''
' Build
' (c) RadiusCore Ltd - https://radiuscore.co.nz/
'
' Build script for RadiusCore Excel addin.
'
' @author Andrew Pullon | andrew.pullon@pkfh.co.nz | andrewcpullon@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

' App information.
Const app_Title = "VBA Scripting"
Const app_Version = "0.2.0"
Const app_Publisher = "RadiusCore Ltd"
Const app_Ext = ".xlam"

' Repo folder locations.
Const rc_SrcFolder = "..\src\"
Const rc_BuildFolder = ".\"
Const rc_RepoRootFolder = "..\"

' Excel add-on Registry Keys.
Const rc_reg_Excel = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Options\OPEN"

' Global objects used to establish dependencies.
Dim rc_Excel
Dim rc_ExcelWasOpen
Dim rc_Workbook
Dim rc_FileSystem
Dim rc_DomDocument
Dim rc_Shell
Dim rc_WscShell
Dim rc_LogFile

Set rc_Excel = Nothing
Set rc_Workbook = Nothing
Set rc_FileSystem = Nothing
Set rc_DomDocument = Nothing
Set rc_Shell = Nothing
Set rc_WscShell = Nothing
Set rc_LogFile = Nothing

' Error handling (return code).
Dim vb_ExitCode
vb_ExitCode = 0

' ============================================= '
' Public Methods
' ============================================= '

Main

Public Sub Main()
	' Variables.
	Dim main_Action
	main_CreateLogFile
	
	PrintLn "---------------------------------------------"
	PrintLn Space(15) & app_Title & " v" & app_Version
	PrintLn Space(17) & "BuildScript"
	PrintLn Space(12) & "www.radiuscore.co.nz"
	PrintLn "---------------------------------------------"
	
	If Not main_Dependencies Then
		PrintLn vbNullString
		LogLn "ERROR Required dependency is unavailable. Exiting BuildScript...", "Main", True
	Else
		If main_Build Then ' Perform build.
			Log "Build sucessfully completed! ", "Main", True: Log vbNewLine, vbNullString, False
		End If
	End If
	
	' Close workbook & Excel (if it wasn't open).
	If Not rc_Workbook Is Nothing Then excel_CloseWorkbook rc_Workbook, vbNullString
	If Not rc_Excel Is Nothing Then excel_CloseExcel rc_Excel, rc_ExcelWasOpen
	WScript.Quit vb_ExitCode
End Sub

' ============================================= '
' MAIN Private Methods
' ============================================= '

''
' Build Excel Addin File. 
''
Private Function main_Build()
	On Error Resume Next
	main_Build = False
	
	' Check for existing build.
	If rc_FileSystem.FileExists(FullPath(rc_RepoRootFolder & app_Title & app_Ext)) Then
		LogLn "WARNING Existing build detected.", "Main.Build", False
		Select Case UCase(Split(Input(vbNewLine & "Warning: Existing build detected. Do you want to continue? (y/n) <")," ")(0)) 
		Case "N"
			PrintLn vbNullString: LogLn "Build cancelled by user. Exiting BuildScript...", "Main.Build", True
			Exit Function
		Case "Y"
			' Do nothing, continue.
			LogLn "User confirmed overwrite of existing build.", "Main.Build", False
			rc_FileSystem.DeleteFile FullPath(rc_RepoRootFolder & app_Title & app_Ext), True
		Case Else
			PrintLn vbNullString: LogLn "ERROR Unrecognized action. Exiting BuildScript...", "Main.Build", True
			Exit Function
		End Select
	End If
	
	' Add Workbook contents to Excel file.
	build_Worksheets
	If Not Err.Number = 0 Then
		PrintLn vbNullString: LogLn "ERROR Failed to import Excel worksheets to Excel document (" & Err.Number & ": " & Err.Description & "). Exiting BuildScript...", "Main.Build", True
		vb_ExitCode = Err.Number
		main_Cleanup Mode, True
		Exit Function
	End If
	
	' Build Visual Basic project in Excel file. 
	build_VbComponents
	If Not Err.Number = 0 Then
		PrintLn vbNullString: LogLn "ERROR Failed to import Visual Basic components to Excel VBA (" & Err.Number & ": " & Err.Description & "). Exiting BuildScript...", "Main.Build", True
		vb_ExitCode = Err.Number
		main_Cleanup Mode, True
		Exit Function
	End If
	build_VbReferences
	If Not Err.Number = 0 Then
		PrintLn vbNullString: LogLn "ERROR Failed to add required Visual Basic references to Excel VBA (" & Err.Number & ": " & Err.Description & "). Exiting BuildScript...", "Main.Build", True
		vb_ExitCode = Err.Number
		main_Cleanup Mode, True
		Exit Function
	End If
	
	' Save Excel File.
	LogLn "Save & close compiled Excel file to disk: " & FullPath(rc_RepoRootFolder & app_Title & app_Ext), "Main.Build", True
	excel_CloseWorkbook rc_Workbook, FullPath(rc_RepoRootFolder & app_Title & app_Ext)
	If Not Err.Number = 0 Then
		PrintLn vbNullString: LogLn "ERROR Failed to save Excel file (" & Err.Number & ": " & Err.Description & "). Exiting BuildScript...", "Main.Build", True
		vb_ExitCode = Err.Number
		main_Cleanup Mode, True
		Exit Function
	End If
	
	' Add RibbonUI to Excel File.
	build_RibbonUI
	If Not Err.Number = 0 Then
		PrintLn vbNullString: LogLn "ERROR Failed to import Excel Ribbon to Excel document (" & Err.Number & ": " & Err.Description & "). Exiting BuildScript...", "Main.Build", True
		vb_ExitCode = Err.Number
		main_Cleanup Mode, True
		Exit Function
	End If
	
	' Register compiled file with Excel.
	build_Register
	If Not Err.Number = 0 Then
		PrintLn vbNullString: LogLn "ERROR Failed to register compiled file with Excel (" & Err.Number & ": " & Err.Description & "). Exiting BuildScript...", "Main.Build", True
		vb_ExitCode = Err.Number
		main_Cleanup Mode, True
		Exit Function
	End If
	
	' Cleanup.
	main_Cleanup Mode, False
	
	' TODO - Set version number if possible?
	' Set file to ReadOnly.
	'If Mode = "RELEASE" Then rc_FileSystem.GetFile(FullPath(rc_RepoRootFolder & app_Title & app_Ext)).Attributes = 1
	main_Build = True
End Function

''
' Create log file, ignore all errors.
''
Private Sub main_CreateLogFile()
	On Error Resume Next
	Dim mn_Temp
	Set mn_Temp = CreateObject("Scripting.FileSystemObject")
	Set rc_LogFile = mn_Temp.OpenTextFile(rc_BuildFolder & "BuildScript.log", 8, -2)
	Set mn_Temp = Nothing
	Log vbNewLine, vbNullString, False
	Log "----------" & app_Title & " v" & app_Version & " BuildScript" &  "----------", vbNullString, False
	Log vbNewLine, vbNullString, False
End Sub

''
' Clean up any files that were created during the build process.
' This will ensure any file remnants left behind due to an error are cleaned up, as 
' well as moving backup files (from protection methods) to a backup folder.
''
Private Sub main_Cleanup(Mode, Failed)
	On Error Resume Next ' Ignore all errors. 
	
	Dim rel_BackupFolder
	Dim rel_Temp
	rel_BackupFolder = FullPath(rc_BuildFolder & "backup") & "\"
	
	With rc_FileSystem
		' Cleanup any files that may be left over from `build_RibbonUI` and `build_Worksheets` methods.
		If .FolderExists(FullPath(rc_RepoRootFolder & "customUI")) Then .DeleteFolder FullPath(rc_RepoRootFolder & "customUI"), True
		If .FolderExists(FullPath(rc_RepoRootFolder & "xl")) Then .DeleteFolder FullPath(rc_RepoRootFolder & "xl"), True
		If .FolderExists(FullPath(rc_RepoRootFolder & "_rels")) Then .DeleteFolder FullPath(rc_RepoRootFolder & "_rels"), True
		If .FileExists(FullPath(rc_RepoRootFolder & "[Content_Types].xml")) Then .DeleteFile FullPath(rc_RepoRootFolder & "[Content_Types].xml"), True
		If .FileExists(FullPath(rc_RepoRootFolder & app_Title & ".zip")) Then .DeleteFile FullPath(rc_RepoRootFolder & app_Title & ".zip"), True
		
		If Mode = "RELEASE" Then
			' Delete/move backup files & log created by `release_Protection` method.
			If .FolderExists(rel_BackupFolder) Then .DeleteFolder Left(rel_BackupFolder, Len(rel_BackupFolder) - 1), True
			.CreateFolder(rel_BackupFolder)
			If .FileExists(FullPath(rc_RepoRootFolder & app_Title & app_Ext & ".cleanbackup")) Then .MoveFile FullPath(rc_RepoRootFolder & app_Title & app_Ext & ".cleanbackup"), rel_BackupFolder & app_Title & app_Ext & ".cln_backup"
			If .FileExists(FullPath(rc_RepoRootFolder & app_Title & app_Ext & ".backup")) Then .MoveFile FullPath(rc_RepoRootFolder & app_Title & app_Ext & ".backup"), rel_BackupFolder & app_Title & app_Ext & ".unv_backup"
			If .FileExists(FullPath(rc_RepoRootFolder & app_Title & app_Ext & ".cco_backup")) Then .MoveFile FullPath(rc_RepoRootFolder & app_Title & app_Ext & ".cco_backup"),   rel_BackupFolder & app_Title & app_Ext & ".cco_backup"
			If .FileExists(FullPath(rc_RepoRootFolder & app_Title & "_e" & app_Ext)) Then
				.MoveFile FullPath(rc_RepoRootFolder & app_Title & app_Ext),  rel_BackupFolder & app_Title & app_Ext & ".enc_backup"
				.MoveFile FullPath(rc_RepoRootFolder & app_Title & "_e" & app_Ext),  FullPath(rc_RepoRootFolder & app_Title & app_Ext)
			End If
			If .FileExists(FullPath(rc_RepoRootFolder & app_Title & ".log")) Then .MoveFile FullPath(rc_RepoRootFolder & app_Title & ".log"), rel_BackupFolder & "UnviewablePlus.log"
			If .FileExists(FullPath(rc_RepoRootFolder & "cco.log")) Then .MoveFile FullPath(rc_RepoRootFolder & "cco.log"), rel_BackupFolder & "CustomCompression.log"
		End If
		If Failed = True Then
			If .FileExists(FullPath(rc_RepoRootFolder & app_Title & app_Ext)) Then .DeleteFile FullPath(rc_RepoRootFolder & app_Title & app_Ext)
		End If
	End With
End Sub

''
' Initialise all global variables, establishing whether required 
' dependencies for build are available.
''
Private Function main_Dependencies()
	On Error Resume Next
	main_Dependencies = False
	LogLn "--> Checking build dependencies", "Main.Dependencies", True
	
	Log Space(3) & "- Excel", "Main.Dependencies", True
	Log ".", vbNullString, True: rc_ExcelWasOpen = excel_OpenExcel(rc_Excel): Log ".", vbNullString, True
	If Not rc_Excel Is Nothing Then
		Log ".", vbNullString, True: excel_CreateWorkbook rc_Workbook
		If Not rc_Workbook Is Nothing Then
			Log "passed", vbNullString, True: Log vbNewLine, vbNullString, True
		Else
			Log ".failed (" & Err.Description & ")", vbNullString, True: Log vbNewLine, vbNullString, True
			Exit Function
		End If
	Else
		Log ".failed (" & Err.Description & ")", vbNullString, True: Log vbNewLine, vbNullString, True
		Exit Function
	End If
	
	Log Space(3) & "- Visual Basic", "Main.Dependencies", True: Log "..", vbNullString, True
	If vba_IsTrusted(rc_Workbook) Then
		Log ".passed", vbNullString, True: Log vbNewLine, vbNullString, True
	Else
		Log ".failed (Access to the VBA project object model is not trusted in Excel.)", vbNullString, True: Log vbNewLine, vbNullString, True
		Exit Function
	End If
		  
	Log Space(3) & "- FileSystemObject", "Main.Dependencies", True
	Log ".", vbNullString, True: Set rc_FileSystem = CreateObject("Scripting.FileSystemObject"): Log ".", vbNullString, True
	If Not rc_FileSystem Is Nothing Then
		Log ".passed", vbNullString, True: Log vbNewLine, vbNullString, True
	Else
		Log ".failed (" & Err.Description & ")", vbNullString, True: Log vbNewLine, vbNullString, True
		Exit Function
	End If
	
	Log Space(3) & "- DOMDocument", "Main.Dependencies", True
	Log ".", vbNullString, True: Set rc_DomDocument = CreateObject("MSXML2.DOMDocument"): Log ".", vbNullString, True
	If Not rc_DomDocument Is Nothing Then
	   	rc_DomDocument.Async = False
		Log ".passed", vbNullString, True: Log vbNewLine, vbNullString, True
	Else
		Log ".failed (" & Err.Description & ")", vbNullString, True: Log vbNewLine, vbNullString, True
		Exit Function
	End If
	
	Log Space(3) & "- Shell", "Main.Dependencies", True
	Log ".", vbNullString, True: Set rc_Shell =  CreateObject("Shell.Application"): Log ".", vbNullString, True
	If Not rc_Shell Is Nothing Then
		Log ".passed", vbNullString, True: Log vbNewLine, vbNullString, True
	Else
		Log ".failed (" & Err.Description & ")", vbNullString, True: Log vbNewLine, vbNullString, True
		Exit Function
	End If
	
	Log Space(3) & "- Wscript Shell", "Main.Dependencies", True
	Log ".", vbNullString, True: Set rc_WscShell = WScript.CreateObject("WScript.Shell"): Log ".", vbNullString, True
	If Not rc_WscShell Is Nothing Then
		Log ".passed", vbNullString, True: Log vbNewLine, vbNullString, True
	Else
		Log ".failed (" & Err.Description & ")", vbNullString, True: Log vbNewLine, vbNullString, True
		Exit Function
	End If
	
	main_Dependencies = True
	LogLn "--> Done", "Main.Dependencies", True
End Function

' ============================================= '
' Build Private Methods
' ============================================= '

''
' Import Visual Basic components to Excel document.
''
Private Sub build_VbComponents()	
	' If 'vbProject' folder does not exist, skip this step.
	If Not rc_FileSystem.FolderExists(FullPath(rc_SrcFolder & "vbProject")) Then Exit Sub
	
	' Name vb project.
	rc_Workbook.VBProject.Name = Replace(app_Title, " ", "_")
	
	' Import content in 'src/vbProject' folder.
	LogLn "--> Import VB Components", "Build.VbComponents", True
	vbcomp_ImportFolder rc_FileSystem.GetFolder(FullPath(rc_SrcFolder & "vbProject"))
	LogLn "--> Done", "Build.VbComponents", True
End Sub
''
' Recursively import folder contents.
''
Private Sub vbcomp_ImportFolder(ByVal FolderSpec)
	Dim build_Folder
	Dim build_File
	Dim build_Ext
	
	For Each build_Folder In FolderSpec.SubFolders
		vbcomp_ImportFolder build_Folder
	Next
	
	For Each build_File in FolderSpec.Files
		build_Ext = rc_FileSystem.GetExtensionName(build_File.Name)
		Select Case build_Ext
		Case "bas", "frm", "cls"
			LogLn Space(3) & "- " & build_File.Name, "Build.VbComponents", True
			vba_ImportModule rc_Workbook, FolderSpec, build_File.Name
		Case "doccls"
			LogLn Space(3) & "- " & build_File.Name, "Build.VbComponents", True
			vba_ImportLines rc_Workbook, FolderSpec, build_File.Name
		End Select
	Next
End Sub

''
' Add necessary references to Visual Basic project.
''
Private Sub build_VbReferences()
	LogLn "--> Add VB References", "Build.VbReferences", True
	LogLn Space(3) & "- NONE", "Build.VbReferences", True
	'rc_Workbook.VBProject.References.AddFromFile {file path here.}
	'rc_Workbook.VBProject.References.AddFromGuid "{GUID here}",0 , 0
	LogLn "--> Done", "Build.VbReferences", True
End Sub

''
' Import Custom Ribbon UI to Excel document.
''
Private Sub build_RibbonUI()
	' Variables.
	Dim build_FolderZip
	Dim build_Element
	
	' If 'customUI' folder does not exist, skip this step.
	If Not rc_FileSystem.FolderExists(FullPath(rc_SrcFolder & "customUI")) Then Exit Sub
	LogLn "--> Import Custom Excel Ribbon UI", "Build.RibbonUI", True
	
	' 1) Change file extension to .zip & get folder object.
	LogLn Space(3) & "- Convert .xlam to .zip", "Build.RibbonUI", True
	rc_FileSystem.MoveFile FullPath(rc_RepoRootFolder & app_Title & app_Ext), FullPath(rc_RepoRootFolder & app_Title & ".zip")
	Set build_FolderZip = rc_Shell.NameSpace(FullPath(rc_RepoRootFolder & app_Title & ".zip"))
	
	' 2) Import CustomUI to zip folder.
	LogLn Space(3) & "- Import customUI folder to .zip", "Build.RibbonUI", True
	rc_FileSystem.CopyFolder FullPath(rc_SrcFolder & "customUI\"), FullPath(rc_RepoRootFolder & "customUI")
	build_FolderZip.MoveHere FullPath(rc_RepoRootFolder & "customUI"), 4
	Do Until Not rc_FileSystem.FolderExists(FullPath(rc_RepoRootFolder & "customUI")): WScript.Sleep(10): Loop ' Wait for copy to finish.
	
	' 3) Extract & edit the .rels file to add relationship to the CustomUI.
	LogLn Space(3) & "- Add document level relationship to Ribbon UI", "Build.RibbonUI", True
	rc_Shell.NameSpace(FullPath(rc_RepoRootFolder)).MoveHere build_FolderZip.Items().Item(1).GetFolder, 4
	Do Until rc_FileSystem.FileExists(FullPath(rc_RepoRootFolder & "_rels\.rels")): WScript.Sleep(10): Loop ' Wait for copy to finish.
    With rc_DomDocument
        .LoadXML rc_FileSystem.OpenTextFile(FullPath(rc_RepoRootFolder & "_rels\.rels"), 1).ReadAll
        Set build_Element = rc_DomDocument.CreateNode(1, "Relationship", .ChildNodes.Item(1).NamespaceURI)
        With build_Element
            .SetAttribute "Id", "R0523d909d7594e1a"
            .SetAttribute "Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
            .SetAttribute "Target", "customUI/customUI14.xml"
        End With
        .ChildNodes.Item(1).AppendChild build_Element
        rc_FileSystem.OpenTextFile(FullPath(rc_RepoRootFolder & "_rels\.rels"), 2).Write .XML 
    End With
	build_FolderZip.MoveHere FullPath(rc_RepoRootFolder & "_rels\"), 4
	Do Until Not rc_FileSystem.FolderExists(FullPath(rc_RepoRootFolder & "_rels")): WScript.Sleep(10): Loop ' Wait for copy to finish.
			
	' 4) Export ContentTypes and edit to allow .png images.
	LogLn Space(3) & "- Add .png content type to default XML structure.", "Build.RibbonUI", True
	rc_Shell.NameSpace(FullPath(rc_RepoRootFolder)).MoveHere build_FolderZip.Items().Item(0), 4
	Do Until rc_FileSystem.FileExists(FullPath(rc_RepoRootFolder & "[Content_Types].xml")): WScript.Sleep(10): Loop ' Wait for copy to finish.
	rc_DomDocument.LoadXML rc_FileSystem.OpenTextFile(FullPath(rc_RepoRootFolder & "[Content_Types].xml"), 1).ReadAll
	xml_contenttype_addDefault rc_DomDocument, "png", "image/png"
	rc_FileSystem.OpenTextFile(FullPath(rc_RepoRootFolder & "[Content_Types].xml"), 2).Write rc_DomDocument.Xml
	build_FolderZip.MoveHere FullPath(rc_RepoRootFolder & "[Content_Types].xml"), 4
	Do Until Not rc_FileSystem.FileExists(FullPath(rc_RepoRootFolder & "[Content_Types].xml")): WScript.Sleep(10): Loop ' Wait for copy to finish.
	
	' 5) Change file extension back to .xl**.
	LogLn Space(3) & "- Convert .zip back to" & app_Ext, "Build.RibbonUI", True
	rc_FileSystem.MoveFile FullPath(rc_RepoRootFolder & app_Title & ".zip"), FullPath(rc_RepoRootFolder & app_Title & app_Ext)
	LogLn "--> Done", "Build.RibbonUI", True
End Sub

''
' Import worksheets to Excel document. 
''
Private Sub build_Worksheets()
	' Variables.
	Dim build_FolderZip
	Dim build_Item
	Dim build_Folder
	Dim build_Element
	Dim build_ContentDefinitions
	Dim build_ItemExt
	
	' Compile dictionary with known content definitions. **THESE WILL NEED UPDATING, AS MORE ARE DISCOVERED**
	Set build_ContentDefinitions = CreateObject("Scripting.Dictionary")
	With build_ContentDefinitions
		.Add "activeX", CreateObject("Scripting.Dictionary")
		.Item("activeX").Add "xml", "application/vnd.ms-office.activeX+xml"
		.Item("activeX").Add "bin", "application/vnd.ms-office.activeX"
		.Add "ctrlProps", CreateObject("Scripting.Dictionary")
		.Item("ctrlProps").Add "xml", "application/vnd.ms-excel.controlproperties+xml"
		.Add "drawings", CreateObject("Scripting.Dictionary")
		.Item("drawings").Add "xml", "application/vnd.openxmlformats-officedocument.drawing+xml"
		.Item("drawings").Add "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"
		.Add "externalLinks", CreateObject("Scripting.Dictionary")
		.Item("externalLinks").Add "xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
		.Add "media", CreateObject("Scripting.Dictionary")
		.Item("media").Add "png", "image/png"
		.Item("media").Add "svg", "image/svg+xml"
		.Item("media").Add "emf", "image/x-emf"
		'***More Media Types??***
		.Add "printerSettings", CreateObject("Scripting.Dictionary")
		.Item("printerSettings").Add "bin", "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
		.Add "tables", CreateObject("Scripting.Dictionary")
		.Item("tables").Add "xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
		.Add "theme", CreateObject("Scripting.Dictionary")
		.Item("theme").Add "xml", "application/vnd.openxmlformats-officedocument.theme+xml"
		.Add "worksheets", CreateObject("Scripting.Dictionary")
		.Item("worksheets").Add "xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	End With
	
	' If 'xl' folder does not exist, skip this step.
	If Not rc_FileSystem.FolderExists(FullPath(rc_SrcFolder & "xl")) Then Exit Sub
	LogLn "--> Import Excel Worksheets", "Build.Worksheets", True
	
	'1) Create blank Excel file on disk.
	LogLn Space(3) & "- Create blank Excel file on disk: " & FullPath(rc_RepoRootFolder & app_Title & app_Ext), "Build.Worksheets", True
	excel_CloseWorkbook rc_Workbook, FullPath(rc_RepoRootFolder & app_Title & app_Ext)
	If Not Err.Number = 0 Then
		Err.Raise Err.Nnumber, ,"Failed to create blank Excel file (" & Err.Description & ")"
		Exit Sub
	End If
	
	' 2) Change file extension to .zip & get folder object.
	LogLn Space(3) & "- Convert " & app_Ext & " to .zip", "Build.Worksheets", True
	rc_FileSystem.MoveFile FullPath(rc_RepoRootFolder & app_Title & app_Ext), FullPath(rc_RepoRootFolder & app_Title & ".zip")
	Set build_FolderZip = rc_Shell.NameSpace(FullPath(rc_RepoRootFolder & app_Title & ".zip"))
	
	' 3) Update contents of 'xl' folder with repo src folder contents.
	' Remove existing 'xl' folder from zip and delete.
	Log Space(3) & "- Update contents of xl folder", "Build.Worksheets", True
	For Each build_Item In build_FolderZip.Items()
		If build_Item.Name = "xl" Then
			rc_Shell.NameSpace(FullPath(rc_RepoRootFolder)).MoveHere build_Item.GetFolder, 4 : Log ".", vbNullString, True
			Exit For
		End If
	Next
	rc_FileSystem.DeleteFolder FullPath(rc_RepoRootFolder & "xl"): Log ".", vbNullString, True
	' Copy 'src\xl' folder to root repo folder. 
	rc_FileSystem.CopyFolder FullPath(rc_SrcFolder & "xl\"), FullPath(rc_RepoRootFolder & "xl"): Log ".", vbNullString, True
	' Remove white space from all .xml documents.
	'xml_RemoveWhiteSpace rc_FileSystem.GetFolder(FullPath(rc_RepoRootFolder & "xl"))
	' Move 'xl' folder into zip.
	build_FolderZip.MoveHere FullPath(rc_RepoRootFolder & "xl"), 4: Log ".", vbNullString, True
	Do Until Not rc_FileSystem.FolderExists(FullPath(rc_RepoRootFolder & "xl")): WScript.Sleep(10): Loop ' Wait for copy to finish.
	Log "complete", vbNullString, True: Log vbNewLine, vbNullString, True
	
	' 4) Export ContentTypes and update to include elements in `xl` folder.
	Log Space(3) & "- Update Excel '[Content_Types].xml' to reflect contents of xl folder", "Build.Worksheets", True
	rc_Shell.NameSpace(FullPath(rc_RepoRootFolder)).MoveHere build_FolderZip.Items().Item(0), 4 ' Export ContentTypes file.
	Do Until rc_FileSystem.FileExists(FullPath(rc_RepoRootFolder & "[Content_Types].xml")): WScript.Sleep(10): Loop ' Wait for copy to finish.
    rc_DomDocument.LoadXML rc_FileSystem.OpenTextFile(FullPath(rc_RepoRootFolder & "[Content_Types].xml"), 1).ReadAll
    ' Update file with new content types.
	xml_contenttype_Validate rc_DomDocument ' Validate raw ContentTypes before changes.
	For Each build_Item In rc_FileSystem.GetFolder(FullPath(rc_SrcFolder & "xl\")).Files
    	If build_Item.Name = "metadata.xml" Then
    		xml_contenttype_addOverride rc_DomDocument, "/xl/" & build_Item.Name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
    	Else
    		xml_contenttype_addOverride rc_DomDocument, "/xl/" & build_Item.Name, "application/vnd.openxmlformats-officedocument.spreadsheetml." & RemoveExtension(build_Item.Name) & "+xml"
    	End If
    Next
	For Each build_Folder In build_ContentDefinitions.Keys
		If rc_FileSystem.FolderExists(FullPath(rc_SrcFolder & "xl\" & build_Folder)) Then
			For Each build_Item In rc_FileSystem.GetFolder(FullPath(rc_SrcFolder & "xl\" & build_Folder)).Files
				build_ItemExt = rc_FileSystem.GetExtensionName(build_Item.Path)
				Select Case build_ItemExt 
				Case "xml", "bin"
					Log ".", vbNullString, True
					xml_contenttype_addOverride rc_DomDocument, "/xl/" & build_Folder & "/" & build_Item.Name, build_ContentDefinitions.Item(build_Folder).Item(build_ItemExt)
				Case "emf", "png", "svg", "vml"
					Log ".", vbNullString, True
					xml_contenttype_addDefault rc_DomDocument, build_ItemExt, build_ContentDefinitions.Item(build_Folder).Item(build_ItemExt)
				Case Else
					Log ".failed (file extension " & build_ItemExt & " not recognised.)", vbNullString, True: Log vbNewLine, vbNullString, True
					Err.Raise 440, , "Unable to build [Content_Types] file due to unrecognised file extension (" & build_ItemExt & ")"
					Exit Sub
				End Select
			Next
		End If
	Next
	' Write modified XML to file & move back into zip.
    xml_contenttype_Validate rc_DomDocument ' Validate raw ContentTypes after changes.
	rc_FileSystem.OpenTextFile(FullPath(rc_RepoRootFolder & "[Content_Types].xml"), 2).Write rc_DomDocument.xml ' Write changes to file.
	build_FolderZip.MoveHere FullPath(rc_RepoRootFolder & "[Content_Types].xml"), 4 ' Move into zip.
	Do Until Not rc_FileSystem.FileExists(FullPath(rc_RepoRootFolder & "[Content_Types].xml")): WScript.Sleep(10): Loop ' Wait for copy to finish.
	Log "complete", vbNullString, True: Log vbNewLine, vbNullString, True
	
	' 5) Change file extension back to .xl**.
	LogLn Space(3) & "- Convert .zip back to" & app_Ext, "Build.Worksheets", True
	rc_FileSystem.MoveFile FullPath(rc_RepoRootFolder & app_Title & ".zip"), FullPath(rc_RepoRootFolder & app_Title & app_Ext)
	WScript.Sleep(100)
	
	' 6) Open Excel document.
	LogLn Space(3) & "- Open Excel document containing imported worksheets", "Build.Worksheets", True
	excel_OpenWorkbook FullPath(rc_RepoRootFolder & app_Title & app_Ext), rc_Workbook
	If rc_Workbook is Nothing Or Not Err.Number = 0 Then
		Err.Raise Err.Nnumber, ,"Failed to open compiled Excel file (" & Err.Description & ")"
		Exit Sub
	End If
	LogLn "--> Done", "Build.Worksheets", True
End Sub

''
' Register compiled Excel file as On-Open Excel Addon.
''
Sub build_Register()
	Select Case UCase(Split(Input(vbNewLine & "Register compiled " & app_Title & app_Ext & " to open with Excel? (y/n) <")," ")(0))
	Case "N"
		PrintLn vbNullString: LogLn "User skipped Excel registration.", "Build.Register", True
		Exit Sub
	Case "Y"
		Dim reg_Item
		Dim reg_TargetKey
		' Determine if file is already registered.
		For Each reg_Item In Split(",1,2,3,4,5,6,7,8,9,10",",")
			If registry_KeyExists(rc_reg_Excel & reg_Item) Then
				If rc_FileSystem.GetFileName(Replace(registry_ReadValue(rc_reg_Excel & reg_Item),"""",vbNullString)) = app_Title & app_Ext Then
					If Replace(registry_ReadValue(rc_reg_Excel & reg_Item),"""",vbNullString) = FullPath(rc_RepoRootFolder & app_Title & app_Ext) Then
						PrintLn vbNullString: LogLn "Compiled file is already registered with Excel.", "Build.Register", True
						Exit Sub
					Else
						' Addin is registered, but not to the correct location. Save key for use later.
						reg_TargetKey = rc_reg_Excel & reg_Item
						Exit For
					End If
				End If
			End If
		Next
		
		' If no TargetKey, find the first available key to register with.
		If reg_TargetKey = vbNullString Then
			For Each reg_Item In Split(",1,2,3,4,5,6,7,8,9,10",",")
				If Not registry_KeyExists(rc_reg_Excel & reg_Item) Then
					reg_TargetKey = rc_reg_Excel & reg_Item
					Exit For
				End If
			Next
		End If
		
		' Register addon to target key.
		If registry_WriteValue(rc_reg_Excel & reg_Item, """" &  FullPath(rc_RepoRootFolder & app_Title & app_Ext) & """", "REG_SZ") Then
			PrintLn vbNullString: LogLn "Sucessfully registered compiled file with Excel.", "Build.Register", True
		Else
			PrintLn vbNullString: LogLn "Failed to register compiled file. You will need to manually do so via Excel.", "Build.Register", True
		End If
	Case Else
		PrintLn vbNullString: LogLn "ERROR Unrecognized action. Registration skipped.", "Build.Register", True
		Exit Sub
	End Select
End Sub

' ============================================= '
' XML Helper Methods
' ============================================= '

''
' Remove whitespace from all `.xml` files in folder (recursive).
''
Sub xml_RemoveWhiteSpace(ByVal TargetFolder)
	Dim xml_Item
	Dim xml_Text
	
	' Recursively trigger function for subfolders.
	For Each xml_Item In TargetFolder.SubFolders
		xml_RemoveWhiteSpace xml_Item
	Next 
	
	' Remove whitespace from files.
	For Each xml_Item In TargetFolder.Files
		If rc_FileSystem.GetExtensionName(xml_Item.Path) = "xml" Then
			Log ".", vbNullString, True
			xml_Text = rc_FileSystem.OpenTextFile(xml_Item.Path, 1).ReadAll
			xml_Text = Replace(Replace(Replace(xml_Text, Chr(13), vbNullString), Chr(10), vbNullString), Chr(9), vbNullString)
			rc_FileSystem.OpenTextFile(xml_Item.Path, 2).Write xml_Text
		End If
	Next
End Sub

''
' Validates contents of the '[Content_Type].xml' file, removing nodes if they are not valid.
''
Sub xml_contenttype_Validate(ByRef XmlDocument)
	Dim xml_Element
	Dim xml_PartName
	For Each xml_Element In XmlDocument.ChildNodes.Item(1).ChildNodes
		Select Case xml_Element.NodeName
		Case "Default"
			' Skip.
		Case "Override"
			' Check if 'PartName' attribute exists, if it begins with '/xl/'.
			xml_PartName = xml_Element.Attributes.getNamedItem("PartName").Text
			If Left(xml_PartName, 4) = "/xl/" Then
				If Not rc_FileSystem.FileExists(FullPath(rc_SrcFolder & xml_PartName)) Then
					Print "removing invalid default node: " & xml_Element.Attributes.getNamedItem("PartName").Text '*** Temp for testing ***
					XmlDocument.ChildNodes.Item(1).RemoveChild xml_Element
				End If
			End If
		Case Else
			PrintLn "Warning: 'xml_contenttype_Validate' method found an unknown node name (" & xml_Element.NodeName & ")."
		End Select
	Next 
End Sub

''
' Add a 'Override' node to '[Content_Type].xml' file, without creating duplicates.
''
Sub xml_contenttype_addOverride(ByRef XmlDocument, ByVal PartName, ByVal ContentType)
	Dim xml_Element
	
	' Detect if 'Override' node already exists.
	For Each xml_Element In XmlDocument.ChildNodes.Item(1).ChildNodes
		If xml_Element.nodeName = "Override" Then
			If xml_Element.Attributes.getNamedItem("PartName").Text = PartName Then
				Exit Sub ' Node exists, exit without adding.
			End If
		End If
	Next
	
	' Add 'Override' node if it doesn't exist.	
	Set xml_Element = XmlDocument.CreateNode(1, "Override", XmlDocument.ChildNodes.Item(1).NamespaceURI)
	With xml_Element
    	.SetAttribute "PartName", PartName
    	.SetAttribute "ContentType", ContentType
	End With
	XmlDocument.ChildNodes.Item(1).AppendChild xml_Element
End Sub

''
' Add a 'Default' node to '[Content_Type].xml' file, without creating duplicates.
''
Sub xml_contenttype_addDefault(ByRef XmlDocument, ByVal Extension, ByVal ContentType)
	Dim xml_Element
	
	' Detect if 'Default' node already exists.
	For Each xml_Element In XmlDocument.ChildNodes.Item(1).ChildNodes
		If xml_Element.nodeName = "Default" Then
			If xml_Element.Attributes.getNamedItem("Extension").Text = Extension Then
				Exit Sub ' Node exists, exit without adding.
			End If
		End If
	Next 
	' Add 'Default' node if it doesn't exist.
	Set xml_Element = XmlDocument.CreateNode(1, "Default", XmlDocument.ChildNodes.Item(1).NamespaceURI)
	With xml_Element
		.SetAttribute "Extension", Extension
		.SetAttribute "ContentType", ContentType
	End With
	XmlDocument.ChildNodes.Item(1).InsertBefore xml_Element, XmlDocument.ChildNodes.Item(1).ChildNodes(1)
End Sub


' ============================================= '
' Registry Helper Methods
' ============================================= '

''
' Determine whether a Key with the given `Value` exists within `SearchKey` and its sub-keys (NOT RECURSIVE).
'
' @param {String} HKey | Root registry identifier, in shorthand.
' @param {String} SearchKey | Key to search for `Value`. Don't include the Root Key in this string.
' @param {String} Value | Value to search for in given `SearchKey`.
' @return {Boolean} | True if Value is found in `SearchKey`, False if not. 
''
Private Function registry_Exists(HKey, SearchKey, Value)
	On Error Resume Next
	Dim reg_HKeyID
	Dim reg_Registry
	Dim reg_KeyList
	Dim reg_Key 
	Dim reg_Temp
	
	Select Case HKey 
	Case "HKCR" 
		reg_HKeyID = &H80000000
	Case "HKCU"
		reg_HKeyID = &H80000001
	Case "HKLM" 
		reg_HKeyID = &H80000002
	Case "HKU"
		reg_HKeyID = &H80000003
	Case "HKCC"
		reg_HKeyID = &H80000005
	Case Else
		Err.Raise 13
	End Select
	If Err.Number = 0 Then
		registry_Exists = False
		Set reg_Registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	
		reg_Registry.EnumKey reg_HKeyID, SearchKey, reg_KeyList
		If Err.Number =0 Then
			For Each reg_Key In reg_KeyList
				If reg_Registry.EnumKey(reg_HKeyID, SearchKey & reg_Key & "\" & Value, reg_Temp) = 0 Then
					registry_Exists = True
					Exit For
				End If
			Next
		Else
			registry_Exists = False
			Err.Clear
		End If
	Else
		registry_Exists = False
		Err.Clear
	End If
End Function

''
' Create and/or update a registry key.
'
' @param {String} TargetKey | Full registry path to Key (if ending in \) or Value-Name to create/update.
' @param {String} Value | Value to set. Only use if TargetKey is a Value-Name.
' @param {KeyType} | Type of Key to create. Only use if TargetKey is a Value-Name.
' @return {Boolean} Whether the registry key was sucessfully created/updated.
''
Private Function registry_WriteValue(TargetKey, Value, KeyType)
	On Error Resume Next
	If KeyType = "REG_SZ" Or KeyType = "REG_DWORD" Or KeyType = "REG_BINARY" Or KeyType = "REG_EXPAND_SZ" Then
		rc_WscShell.RegWrite TargetKey, Value, KeyType	
	Else
		Err.Raise 13
	End If
	
	If Err.Number = 0 then
		registry_WriteValue = True
	Else
		registry_WriteValue = False
		Err.Clear
	End If
End Function

''
' Determine whether given `TargetKey` exists in registry.
'
' @param {String} TargetKey | Full registry path to Key or Value-Name.
' @return {Boolean}
''
Private Function registry_KeyExists(TargetKey)
	On Error Resume Next
	Dim Temp
	Temp = rc_WscShell.RegRead(TargetKey)
	If Not Err.Number = 0 then
		If Right(TargetKey,1)="\" Then    ' TargetKey is a registry key-name.
			If InStr(1, Err.Description, "Unable to open registry key", 1) <> 0 Then
				registry_KeyExists = True
			Else
				registry_KeyExists = False
			End If
		Else    ' `TargetKey` is a registry value-name
			registry_KeyExists = False
		End If
		Err.Clear
	Else
		registry_KeyExists = True
	End If
End Function

''
' Read value of `TargetValue` from the registry.
'
' @param {String} TargetValue | Full registry path to Value-Name.
' @return {String} | Returns vbNullString on error, else value of registry key.
''
Private Function registry_ReadValue(TargetValue)
      On Error Resume Next
	  
	  Dim reg_KeyValue
	  reg_KeyValue = rc_WscShell.RegRead(TargetValue)
	  
	  If Err.Number = 0 Then
		registry_ReadValue = reg_KeyValue
      Else
		registry_ReadValue = vbNullString
		Err.Clear
      End If
End Function

' ============================================= '
' Excel Helper Methods
' ============================================= '

''
' Create new Excel document & set some properties.
'
' @param {Object} Workbook object to load Workbook into
''
Private Function excel_CreateWorkbook(ByRef Workbook)
	Set Workbook = rc_Excel.Workbooks.Add()
	If app_Ext = ".xlam" Then Workbook.IsAddin = True
	With Workbook.BuiltinDocumentProperties
		.Item("Company") = app_Publisher
		.Item("Title") = app_Title
		.Item("Author") = app_Publisher
		.Item("Subject") = "Connects Excel with Xero online services."
	End With
End Function

''
' Open existing workbook.
'
' @param {Object} Workbook object to load Workbook into
''
Private Function excel_OpenWorkbook(Path, ByRef Workbook)
	If rc_FileSystem.FileExists(FullPath(Path)) Then
		Set Workbook = rc_Excel.Workbooks.Open(FullPath(Path))
	Else
		Set Workbook = Nothing
	End If
End Function

''
' Close Workbook. Save to Path if specified.
'
' @param {Object} Workbook | Target Workbook to close.
' @param {String} Path | Path to save workbook to before closing. Pass vbNullString to close without saving.
''
Private Sub excel_CloseWorkbook(ByRef Workbook, Path)
	If Not Workbook Is Nothing Then
		If Not Path = vbNullString Then
			If rc_FileSystem.FileExists(Path) Then
				' Existing workbook, use 'Save'.
				Workbook.Save
			Else
				' New workbook, use 'SaveAs' to specify path & file type.
				Select Case app_Ext
				Case ".xlam"
					Workbook.SaveAs Path, 55
				Case ".xlsm"
					Workbook.SaveAs Path, 52
				Case Else
					Workbook.SaveAs Path
				End Select
			End If
		End If
		Workbook.Close False
	End If

	Set Workbook = Nothing
End Sub

''
' Open Excel and return whether Excel was already open.
'
' @param {Object} Excel object to load Excel into
' @return {Boolean} Excel was already open
''
Private Function excel_OpenExcel(ByRef Excel)
	On Error Resume Next
	
	Set Excel = GetObject(, "Excel.Application")

	If Excel Is Nothing Or Not Err.Number = 0 Then
		Err.Clear
		Set Excel = CreateObject("Excel.Application")
		excel_OpenExcel = False
	Else
		excel_OpenExcel = True
	End If
End Function

''
' Close Excel (keep open if previously open).
'
' @param {Object} Excel
' @param {Boolean} KeepExcelOpen
''
Private Sub excel_CloseExcel(ByRef Excel, KeepExcelOpen)
	If Not KeepExcelOpen And Not Excel Is Nothing Then
		Excel.Quit  
	End If
	
	Set Excel = Nothing
End Sub

' ============================================= '
' VBA Helper Methods
' ============================================= '

''
' Check if VBA is trusted
'
' @param {Object} Workbook
' @return {Boolean}
''
Private Function vba_IsTrusted(Workbook)
	On Error Resume Next
	Dim xl_Temp
	xl_Temp = Workbook.VBProject.VBComponents.Count

	If Err.Number = 0 Then
		vba_IsTrusted = True
	Else
		Err.Clear
		vba_IsTrusted = False
	End If
End Function

''
' Get module from Workbook VBProject.
'
' @param {Object} Workbook
' @param {String} Name
''
Private Function vba_GetModule(Workbook, Name)
  Dim Module
  Set vba_GetModule = Nothing

  For Each Module In Workbook.VBProject.VBComponents
    If Module.Name = Name Then
      Set vba_GetModule = Module
      Exit Function
    End If
  Next
End Function

''
' Import module (.bas, .cls, .frm) to Workbook.
'
' @param {Object} Workbook | Workbook to import code to. 
' @param {String} Folder | Source folder. 
' @param {String} Filename | Source file name.
''
Private Sub vba_ImportModule(Workbook, Folder, Filename)
  Dim v_Module
  If Not Workbook Is Nothing Then
    ' Check for existing and remove
    Set v_Module = vba_GetModule(Workbook, RemoveExtension(Filename))
    If Not v_Module Is Nothing Then
      Workbook.VBProject.VBComponents.Remove v_Module
    End If

    ' Import module
    Workbook.VBProject.VBComponents.Import FullPath(rc_FileSystem.BuildPath(Folder, Filename))
	
	' Delete first line if it is blank (happens when importing .frm).
	Set v_Module = vba_GetModule(Workbook, RemoveExtension(Filename))
	If v_Module.CodeModule.Lines(1, 1)="" and v_Module.CodeModule.CountOfLines > 1 Then
		v_Module.CodeModule.DeleteLines 1, 1
	End If
  End If
End Sub

''
' Import code to worksheet objects.
'
' @param {Object} Workbook | Workbook to import code to. 
' @param {String} Folder | Source folder. 
' @param {String} Filename | Source file name.
''
Private Sub vba_ImportLines(Workbook, Folder, Filename)
	Dim v_Module
	If Not Workbook Is Nothing Then
		' Check for existing sheet & add if not present.
		Set v_Module = vba_GetModule(Workbook, RemoveExtension(Filename))
		If v_Module Is Nothing Then
			Dim v_Sheet
			Set v_Sheet = Workbook.Sheets.Add
			Set v_Module = vba_GetModule(Workbook, v_Sheet.Name)
			v_Module.Name = RemoveExtension(Filename)
		End If
		' Import lines.
		v_Module.CodeModule.DeleteLines 1, v_Module.CodeModule.CountOfLines
		v_Module.CodeModule.AddFromFile FullPath(rc_FileSystem.BuildPath(Folder, Filename)) 
	 End If
End Sub

' ============================================= '
' FileSystemObject Helper Methods
' ============================================= '

Private Function FullPath(Path)
  FullPath = rc_FileSystem.GetAbsolutePathName(Path)
End Function

Private Function RemoveExtension(Name)
    Dim Parts
    Parts = Split(Name, ".")
    
    If UBound(Parts) > LBound(Parts) Then
        ReDim Preserve Parts(UBound(Parts) - 1)
    End If
    
    RemoveExtension = Join(Parts, ".")
End Function

' ============================================= '
' General Helper Methods
' ============================================= '

Private Function execStdOut(cmd)
	On Error Resume Next
	Dim exe_App
	Set exe_App = rc_WscShell.Exec(cmd)
	Do While Not exe_App.Status = 1: Wscript.Sleep(10): Loop
	
	Select Case exe_App.Status
	Case 1
		execStdOut = exe_App.StdOut.ReadAll()
	Case 2
		Err.Raise 440, , exe_App.StdErr.ReadAll()
	End Select
End Function 

Private Sub Log(Message, Source, ToPrint)
	If Not rc_LogFile Is Nothing Then
		If Source = vbNullString Then
			rc_LogFile.Write Message
		Else
			Dim lg_DateStamp
			Dim lg_TimeString
			lg_DateStamp = Now
			lg_TimeString = Year(lg_DateStamp) & "-" & Right("0" & Month(lg_DateStamp), 2) & "-" & Right("0" & Day(lg_DateStamp), 2) & " " & _
		    Right("0" & Hour(lg_DateStamp), 2) & ":" & Right("0" & Minute(lg_DateStamp), 2) & ":" & Right("0" & Minute(lg_DateStamp), 2) & "." & Left(Split(Timer & ".0",".")(1) & "00",2)
			
			rc_LogFile.Write lg_TimeString & "|" & Source & "|" & Message
		End If
	End If
	If ToPrint Then Print Message
End Sub

Private Sub LogLn(Message, Source, ToPrint)
	If Not rc_LogFile Is Nothing Then
		Dim lg_DateStamp
		Dim lg_TimeString
		lg_DateStamp = Now
		lg_TimeString = Year(lg_DateStamp) & "-" & Right("0" & Month(lg_DateStamp), 2) & "-" & Right("0" & Day(lg_DateStamp), 2) & " " & _
		Right("0" & Hour(lg_DateStamp), 2) & ":" & Right("0" & Minute(lg_DateStamp), 2) & ":" & Right("0" & Minute(lg_DateStamp), 2) & "." & Left(Split(Timer & ".0",".")(1) & "00",2)
			
		rc_LogFile.WriteLine lg_TimeString & "|" & Source & "|" & Message
	End If
	If ToPrint Then PrintLn Message
End Sub

Private Sub Print(Message)
  WScript.StdOut.Write Message
End Sub

Private Sub PrintLn(Message)
  WScript.StdOut.Write Message & vbNewLine
End Sub

Private Function Input(Prompt)
  If Prompt <> "" Then
    Print Prompt & " "
  End If

  Input = WScript.StdIn.ReadLine 
End Function
