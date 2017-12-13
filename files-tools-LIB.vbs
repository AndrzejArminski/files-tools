Option Explicit

' TODO: Implement data entry syntax:
' copy|move F:\OneDrive\AstroProjects\test\SKY\<AnyFolder>\03\Cas?sss^ttt.txt F:\OneDrive\AstroProjects\test\SKY\<StarNameFolder>\01\
' delete|rename F:\OneDrive\AstroProjects\test\SKY\<AnyFolder>\01\stars?sss^ttt.txt   <--- allow Suffix "01\"  !
' delete F:\OneDrive\AstroProjects\test\SKY\<EmptyFolder>\


' files-tools-LIB.vbs
' Library for copying, moving, deleting and renaming multiple files.
'-------------------------------------------------------------------
' (c)2017-12-12, Andrzej Armiñski, e-mail: andrzejarminski@gmail.com

' Edit & Run *.WSF files to perform these tasks: 

' copy-files-folder-to-folder.WSF
' copy-files-subfolders-to-folder.WSF
' copy-files-folder-to-subfolders.WSF
' copy-files-subfolders-to-subfolders.WSF
' move-files-folder-to-folder.WSF
' move-files-subfolders-to-folder.WSF
' move-files-folder-to-subfolders.WSF
' move-files-subfolders-to-subfolders.WSF
' delete-files-from-folder.WSF
' delete-files-from-subfolders.WSF
' delete-empty-subfolders.WSF
' rename-files-in-folder.WSF
' rename-files-in-subfolders.WSF
' count-files-in-folder.WSF
' count-files-in-subfolders.WSF

' Description of Sub's arguments:
' -------------------------------
' SrcFolder			- Source folder
' DstFolder 		- Destination folder
' SrcRootFolder		- Root folder for source subolders
' SrcFolderSuffix	- "01\", example: "F:\SKY\AA And\01\" (RootFolder\AnySubfolder\FolderSuffix\)
' 						By supplying non-empty DstFolderSuffix you copy/move files from SrcRootFolder's subfolder's subfolder.
' DstRootFolder 	- Root folder for destination subolders
' DstFolderSuffix	- "01\", example: "F:\SKY\AA And\01\" (RootFolder\StarName\FolderSuffix\)
' 						By supplying non-empty DstFolderSuffix you copy/move files to DstRootFolder's StarName-subfolder's subfolder.
' DstFilenamePrefix	- DstFilenamePrefix is attached at the begining of destination filenames. Example: "S_" , "S_7865-4565.fit"
' Extension			- Only files with this extension will be moved or copied
' SearchStr 		- Only files with SearchStr within filename will be moved or copied
' oldStr			- Replace non-empty OldStr with NewStr in filenames while moving or copying files
' newStr
' overWrite			- when True, files in destination will be overwritten

' Global:
	Const ForWriting = 2
	Dim oFSO, logfile, dt, logFN
	Dim count, moved_copied, replaced, skipped, alreadyExists
	Dim first
	count = 0			' Number of SrcFolder(s) files with Extension
	moved_copied = 0	' Number of moved or copied files
	replaced = 0		' Number of overwritten files
	skipped = 0			' Number of skiped files because filename does not contain SearchStr
	alreadyExists = 0	' Number of files that were not moved as already existed in DstFolder
	first = True

Function StarName(FN, Extension)
' Extracts star name from last token of file name FN
' File name format: "SFDB_2016-02-03_1917-29_J000703_BMAH__V_1x1_0040s_HIP 110893.fit"
' File name format: "6789-3447_AA And.xlsx"
' Returns star name: "AA And"
	Dim t
	t = Split(Left(FN, Len(FN) - (Len(Extension) + 1)), "_") 
	StarName = t(Ubound(t)) & "\"
End Function

Function FolderNameWithSlash(FolderName)
	FolderNameWithSlash = FolderName
	If Len (FolderName) > 0 Then
		If Right(FolderName, 1) <> "\" Then FolderNameWithSlash = FolderName & "\"
	End If
End Function

Function FolderExists(Folder)
	If Not oFSO.FolderExists(Folder) Then
		If vbYes = MsgBox("Folder " & Folder & " does not exist." & vbNewLine & "Do you want to create one?", vbYesNo, "Create Folder") Then
			logfile.WriteLine "create folder: " & Folder
			oFSO.CreateFolder Folder
		Else
			WScript.Echo "Script aborted, no file changed."
			WScript.Quit
		End If
	End If
	FolderExists = True
End Function

Sub OpenLogFile(Folder)
	Dim dt, created
	created = False
	If Not oFSO.FolderExists(Folder) Then
		If vbYes = MsgBox("Folder " & Folder & " does not exist." & vbNewLine & "Do you want to create one?", vbYesNo, "Create Folder") Then
			Err.Clear
			On Error Resume Next
			oFSO.CreateFolder Folder
			If Err.Number <> 0 Then
				MsgBox "Incorrect folder path: " & Folder & vbNewLine & "Check spelling." & vbNewLine, vbOK, "ERROR: " & Err.Description
				WScript.Quit 
			End If
			On Error GoTo 0
			created = True
		Else
			WScript.Echo "Script aborted, no file changed."
			WScript.Quit
		End If
	End If
	dt = Now()
	logFN = Folder & "files-tools-logfile-" & Year(dt) & "-" & Month(dt) & "-" & Day(dt) & "-" & Hour(dt) & "-" & Minute(dt) & "-" & Second(dt) & ".txt"
	Set logfile = oFSO.OpenTextFile(logFN, ForWriting, True)
	logfile.WriteLine "#LOGFILE=" & logFN & vbNewLine
	If created Then logfile.WriteLine "create folder: " & Folder
End Sub

Sub WelcomeMessage(Moving, FromSingleFolder, ToSingleFolder, _
					SrcFolder, SrcFolderSuffix, DstFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Input: Moving, FromSingleFolder, ToSingleFolder As Boolean
	Dim title, s, Renaming
	
	OpenLogFile DstFolder
	
	Renaming = Moving And (FromSingleFolder = ToSingleFolder) And (SrcFolder = DstFolder)
	If Renaming Then
		s = "Renaming *."
		title = "Renaming Files "
	Else
		If Moving Then 
			s = "Moving *."
			title = "Moving Files: "
		Else 
			s = "Copying *."
			title = "Copying Files: "
		End If
	End If
	s = s & Extension & " files "
	If Len(SearchStr) > 0 Then s = s + "with string " & """" & SearchStr & """" & " in the file name "
	
	If Renaming Then
		If FromSingleFolder Then
			s = s & vbNewLine & vbNewLine & "in folder " & SrcFolder
			title = title & "In Folder "
		Else
			s = s & vbNewLine & vbNewLine & "in subfolders of folder " & SrcFolder
			title = title & "In Subfolders"
		End If
	Else
		If FromSingleFolder Then
			s = s & vbNewLine & vbNewLine & "from folder " & SrcFolder
			title = title & "Folder To "
		Else
			s = s & vbNewLine & vbNewLine & "from subfolders: " & SrcFolder & "<AnyFolder>\" & SrcFolderSuffix
			title = title & "Subfolders To "
		End If
		If ToSingleFolder Then
			s = s & vbNewLine & vbNewLine & "to folder " & DstFolder 
			title = title & "Folder"
		Else
			s = s & vbNewLine & vbNewLine & "to subfolders: " & DstFolder & "<StarNameFolder>\" & DstFolderSuffix
			title = title & "Subfolders"
		End If
	End If
	
	If Len(DstFilenamePrefix) > 0 Then _
		s = s & vbNewLine & vbNewLine & "The substring "  & """" & DstFilenamePrefix & """" & " to be attached at the begining of destination file names."
	If Len(oldStr) > 0 Then s = s & vbNewLine & vbNewLine & "The substring " & """" & oldStr & """" & " to be replaced with substring " & """" & newStr & """" & " in destination file names."
	
	logfile.WriteLine s
	s = s & vbNewLine & vbNewLine & "Do you want to proceed?"
	
	If vbNo = MsgBox(s, vbYesNo, title) Then 
		MsgBox "Script aborted by user, no file changed.", vbOKOnly, "Script Aborted"
		WScript.Quit
	End If
	
	If overWrite Then
		If vbNo = MsgBox("WARNING: files in destination will be overwritten!" & vbNewLine & "Do you realy want to proceed?", vbYesNo, "WARNING") Then 
			MsgBox "Script aborted by user, no file.", vbOKOnly, "Script Aborted"
			WScript.Quit
		End If
	End If
End Sub

Sub Summary(Moving, FromSingleFolder, ToSingleFolder, SrcFolder, DstFolder, Extension, SearchStr)
	Dim s, Renaming
	
	Renaming = Moving And (FromSingleFolder = ToSingleFolder) And (SrcFolder = DstFolder)

	s = "All done:" & vbNewLine & vbNewLine & count & " *." & Extension & " files were found in "
	If FromSingleFolder Then
		s = s &	"source folder "
	Else
		s = s &	"subfolders of source folder " 
	End If
	s = s & SrcFolder
	
	If moved_copied > 0 Then
		s = s & vbNewLine & vbNewLine & moved_copied & " *." & Extension & " files were "
		If Renaming Then
			s = s & "renamed in "
			If Not ToSingleFolder Then s = s & "subfolders of "
		Else
			If Moving Then
				s = s & "moved to "
			Else
				s = s & "copied to "
			End If
			If ToSingleFolder Then
				s = s & "destination folder " 
			Else
				s = s & "subfolders of destination folder " 
			End If
		End If
		s = s & DstFolder
	End If
	
	If replaced > 0 Then 
		s = s & vbNewLine & vbNewLine & replaced & " *." & Extension & " files were replaced in "
		If ToSingleFolder Then
			s = s & "destination folder " 
		Else
			s = s & "subfolders of destination folder " 
		End If	
		s = s & DstFolder
	End If
	
	If alreadyExists > 0 Then s = s & vbNewLine & vbNewLine & alreadyExists & " *." & Extension & " files were NOT copied/moved/renamed as already existed in destination folder(s)"
	If skipped > 0 Then s = s & vbNewLine & vbNewLine & skipped & " *." & Extension & " files were skipped, as did not contain string: "  & """" & SearchStr & """"
	logfile.WriteLine vbNewLine & s
	logfile.Close
	Set logfile = Nothing
	MsgBox s, vbOKOnly, "Summary"
End Sub

' copy-files-folder-to-folder
Sub CopyFilesFolderToFolder(SrcFolder, DstFolder, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Copies files with Extension and SearchStr in filenames from SrcFolder to DstFolder.
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcFolder = FolderNameWithSlash(SrcFolder)
	DstFolder = FolderNameWithSlash(DstFolder)
	WelcomeMessage False, True, True, SrcFolder, "", DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcFolder) Then
		CopyFiles SrcFolder, DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, False
		Summary False, True, True, SrcFolder, DstFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcFolder
	End If
	Set oFSO = Nothing
End Sub

' copy-files-subfolders-to-folder
Sub CopyFilesSubfoldersToFolder(SrcRootFolder, SrcFolderSuffix, DstFolder, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Copies files with Extension and SearchStr in filenames from SrcRootFolder's subfolders to DstFolder.
	Dim d
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcRootFolder = FolderNameWithSlash(SrcRootFolder)
	SrcFolderSuffix = FolderNameWithSlash(SrcFolderSuffix)
	DstFolder = FolderNameWithSlash(DstFolder)
	WelcomeMessage False, False, True, SrcRootFolder, SrcFolderSuffix, DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcRootFolder) Then
		For Each d In oFSO.GetFolder(SrcRootFolder).SubFolders
			If oFSO.FolderExists(d.Path & "\" & SrcFolderSuffix) Then
				CopyFiles d.Path & "\" & SrcFolderSuffix, DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, False
			Else
				logfile.WriteLine "Source folder does not exist: " & d.Path & "\" & SrcFolderSuffix
			End If
		Next
		Summary False, False, True, SrcRootFolder, DstFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcRootFolder
	End If
	Set oFSO = Nothing
End Sub

' copy-files-folder-to-subfolders
Sub CopyFilesFolderToSubfolders(SrcFolder, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Copies files with Extension and SearchStr in filenames from SrcFolder to DstRootFolder's subfolders using StarNames as subfolder's names.
	Dim d
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcFolder = FolderNameWithSlash(SrcFolder)
	DstRootFolder = FolderNameWithSlash(DstRootFolder)
	DstFolderSuffix = FolderNameWithSlash(DstFolderSuffix)
	WelcomeMessage False, True, False, SrcFolder, "", DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcFolder) Then
		CopyFiles SrcFolder, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, True
		Summary False, True, False, SrcFolder, DstRootFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcFolder
	End If
	Set oFSO = Nothing
End Sub

' copy-files-subfolders-to-subfolders
Sub CopyFilesSubfoldersToSubfolders(SrcRootFolder, SrcFolderSuffix, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Copies files with Extension and SearchStr in filenames from SrcRootFolder's subfolders to DstRootFolder's subfolders using StarNames as subfolder's names.
	Dim d
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcRootFolder = FolderNameWithSlash(SrcRootFolder)
	SrcFolderSuffix = FolderNameWithSlash(SrcFolderSuffix)
	DstRootFolder = FolderNameWithSlash(DstRootFolder)
	DstFolderSuffix = FolderNameWithSlash(DstFolderSuffix)
	WelcomeMessage False, False, False, SrcRootFolder, SrcFolderSuffix, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcRootFolder) Then
		For Each d In oFSO.GetFolder(SrcRootFolder).SubFolders
			If oFSO.FolderExists(d.Path & "\" & SrcFolderSuffix) Then
				CopyFiles d.Path & "\" & SrcFolderSuffix, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, True
			Else
				logfile.WriteLine "Source folder does not exist: " & d.Path & "\" & SrcFolderSuffix
			End If
		Next
		Summary False, False, False, SrcRootFolder, DstRootFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcRootFolder
	End If
	Set oFSO = Nothing
End Sub

' move-files-folder-to-folder
' rename-files-in-folder
Sub MoveFilesFolderToFolder(SrcFolder, DstFolder, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Moves files with Extension and SearchStr in filenames from SrcFolder to DstFolder.
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcFolder = FolderNameWithSlash(SrcFolder)
	SrcFolderSuffix = FolderNameWithSlash(SrcFolderSuffix)
	DstFolder = FolderNameWithSlash(DstFolder)
	WelcomeMessage True, True, True, SrcFolder, "", DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcFolder) Then
		MoveFiles SrcFolder, DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, False
		Summary True, True, True, SrcFolder, DstFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcFolder
	End If
	Set oFSO = Nothing
End Sub

' move-files-subfolders-to-folder
Sub MoveFilesSubfoldersToFolder(SrcRootFolder, SrcFolderSuffix, DstFolder, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Moves files with Extension and SearchStr in filenames from SrcRootFolder's subfolders to DstFolder.
	Dim d
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcRootFolder = FolderNameWithSlash(SrcRootFolder)
	SrcFolderSuffix = FolderNameWithSlash(SrcFolderSuffix)
	DstFolder = FolderNameWithSlash(DstFolder)
	WelcomeMessage True, False, True, SrcRootFolder, SrcFolderSuffix, DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcRootFolder) Then
		For Each d In oFSO.GetFolder(SrcRootFolder).SubFolders
			If oFSO.FolderExists(d.Path & "\" & SrcFolderSuffix) Then
				MoveFiles d.Path & "\" & SrcFolderSuffix, DstFolder, "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, False
			Else
				logfile.WriteLine "Source folder does not exist: " & d.Path & "\" & SrcFolderSuffix
			End If
		Next
		Summary True, False, True, SrcRootFolder, DstFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcRootFolder
	End If
	Set oFSO = Nothing
End Sub

' move-files-folder-to-subfolders
Sub MoveFilesFolderToSubfolders(SrcFolder, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Moves files with Extension and SearchStr in filenames from SrcFolder to DstRootFolder's subfolders using StarNames as subfolder's names.
	Dim d
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcFolder = FolderNameWithSlash(SrcFolder)
	DstRootFolder = FolderNameWithSlash(DstRootFolder)
	DstFolderSuffix = FolderNameWithSlash(DstFolderSuffix)
	WelcomeMessage True, True, False, SrcFolder, "", DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcFolder) Then
		MoveFiles SrcFolder, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, True
		Summary True, True, False, SrcFolder, DstRootFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcFolder
	End If
	Set oFSO = Nothing
End Sub

' move-files-subfolders-to-subfolders
' rename-files-in-subfolders
Sub MoveFilesSubfoldersToSubfolders(SrcRootFolder, SrcFolderSuffix, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite)
' Moves files with Extension and SearchStr in filenames from SrcRootFolder's subfolders to DstRootFolder's subfolders using StarNames as subfolder's names.
	Dim d, Rename
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	SrcRootFolder = FolderNameWithSlash(SrcRootFolder)
	SrcFolderSuffix = FolderNameWithSlash(SrcFolderSuffix)
	DstRootFolder = FolderNameWithSlash(DstRootFolder)
	DstFolderSuffix = FolderNameWithSlash(DstFolderSuffix)
	Rename = (SrcRootFolder = DstRootFolder) And Len(SrcFolderSuffix) = 0 And Len(DstFolderSuffix) = 0
	WelcomeMessage True, False, False, SrcRootFolder, SrcFolderSuffix, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite
	If oFSO.FolderExists(SrcRootFolder) Then
		For Each d In oFSO.GetFolder(SrcRootFolder).SubFolders
			If Rename Then 
				MoveFiles d.Path & "\", d.Path & "\", "", DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, False
			Else
				If oFSO.FolderExists(d.Path & "\" & SrcFolderSuffix) Then
					MoveFiles d.Path & "\" & SrcFolderSuffix, DstRootFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, True
				Else
					logfile.WriteLine " Source folder does not exists: " & d.Path & "\" & SrcFolderSuffix
				End If
			End If
		Next
		Summary True, False, False, SrcRootFolder, DstRootFolder, Extension, SearchStr
	Else
		WScript.Echo "Source folder does not exist: " & SrcRootFolder
	End If
	Set oFSO = Nothing
End Sub

' delete-files-from-folder
' count-files-in-folder
Sub DeleteFilesFromFolder(Folder, Extension, SearchStr, DoDelete)
' Deletes or counts files with Extension and SearchStr in filenames, in Folder.
	Dim f, FN, ss, count, str, action
	Folder = FolderNameWithSlash(Folder)
	If DoDelete Then
		action = "delete: "
		str = "Files in folder " & Folder & vbNewLine & "with extension ." & Extension
		If Len(SearchStr) > 0 Then str = str & vbNewLine & "and containing string " & """" & SearchStr & """"
		str = str & vbNewLine & vbNewLine & "will be permanently deleted!" & vbNewLine & vbNewLine & "Do you realy want to proceed?" 
		If vbNo = MsgBox(str, vbYesNo, "Delete Files") Then
			MsgBox "Script aborted by user, nothing deleted.", vbOKOnly, "Script Aborted"
			WScript.Quit
		End If
	Else
		action = "count: "
		str = "Count files in folder " & Folder & vbNewLine & "with extension ." & Extension
		If Len(SearchStr) > 0 Then str = str & vbNewLine & "and containing string " & """" & SearchStr & """" & "."
		MsgBox str, vbOK, "Count Files"
	End If
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FolderExists(Folder) Then
		str = ""
		count = 0
		For Each f In oFSO.GetFolder(Folder).Files
			If StrComp(oFSO.GetExtensionName(f), Extension) = 0 Then
				FN = oFSO.GetFileName(f)
				ss = Len(SearchStr) = 0 Or InStr(FN, SearchStr) > 0
				If ss Then
					str = str & action & Folder & FN & vbNewLine
					If DoDelete Then oFSO.DeleteFile Folder & FN
					count = count + 1
				End If
			End If
		Next
		OpenLogFile Folder
		logfile.WriteLine str
		If DoDelete Then
			str = vbNewLine & "All done:" & vbNewLine & vbNewLine & count & " files deleted from folder " & Folder
		Else
			str = vbNewLine & "All done:" & vbNewLine & vbNewLine & "There are " & count & " matching files in folder " & Folder
		End If
		logfile.WriteLine str
		logfile.Close
		Set logfile = Nothing
		WScript.Echo str
	Else
		WScript.Echo "Source folder does not exist: " & Folder
	End If
	Set oFSO = Nothing
End Sub

' delete-files-from-subfolders
' count-files-in-subfolders
Sub DeleteFilesFromSubfolders(RootFolder, Extension, SearchStr, DoDelete)
' Deletes or counts files with Extension and SearchStr in filenames, in all subfolders of RootFolder.
	Dim d, f, Folder, FN, ss, count, str, action
	RootFolder = FolderNameWithSlash(RootFolder)
	If DoDelete Then
		action = "delete: "
		str = "Files in all subfolders of folder " & RootFolder & vbNewLine & "with extension ." & Extension
		If Len(SearchStr) > 0 Then str = str & vbNewLine & "and containing string " & """" & SearchStr & """"
		str = str & vbNewLine & vbNewLine & "will be permanently deleted!" & vbNewLine & vbNewLine & "Do you realy want to proceed?" 
		If vbNo = MsgBox(str, vbYesNo, "Delete Files") Then
			MsgBox "Script aborted by user, nothing deleted.", vbOKOnly, "Script Aborted"
			WScript.Quit
		End If
	Else
		action = "count: "
		str = "Count files in all subfolders of folder " & RootFolder & vbNewLine & "with extension ." & Extension
		If Len(SearchStr) > 0 Then str = str & vbNewLine & "and containing string " & """" & SearchStr & """" & "."
		MsgBox str, vbOK, "Count Files"
	End If

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FolderExists(RootFolder) Then
		str = ""
		count = 0
		For Each d In oFSO.GetFolder(RootFolder).SubFolders
			Folder = RootFolder & d.Name & "\"
			For Each f In d.Files
				If StrComp(oFSO.GetExtensionName(f), Extension) = 0 Then
					FN = oFSO.GetFileName(f)
					ss = Len(SearchStr) = 0 Or InStr(FN, SearchStr) > 0
					If ss Then
						str = str & action & Folder & FN & vbNewLine
						If DoDelete Then oFSO.DeleteFile Folder & FN
						count = count + 1
					End If
				End If
			Next
		Next
		OpenLogFile RootFolder
		logfile.WriteLine str
		If DoDelete Then
			str = vbNewLine & "All done:" & vbNewLine & vbNewLine & count & " files deleted from subfolders of " & RootFolder
		Else
			str = vbNewLine & "All done:" & vbNewLine & vbNewLine & "There are " & count & " matching files in subfolders of " & RootFolder
		End If
		logfile.WriteLine str
		logfile.Close
		Set logfile = Nothing
		WScript.Echo str
	Else
		WScript.Echo "Source folder does not exist: " & Folder
	End If
	Set oFSO = Nothing
End Sub

Function FirstOperation(Moving, src, dst, Rename)
	Dim s
	If Rename Then
		s = "Renaming first file"
	Else
		If Moving Then 
			s = "Moving first file" 
		Else
			s = "Copying first file"
		End If
	End If
	s = s & vbNewLine & vbNewLine & "from " & src & vbNewLine & vbNewLine & "to " & dst & vbNewLine & vbNewLine & "Do you want to proceed with all eligible files?"
	If vbNo = MsgBox(s, vbYesNo, "First operation") Then 
		MsgBox "Script aborted by user, no files changed.", vbOKOnly, "Script Aborted"
		WScript.Quit
	End If
	FirstOperation = False
End Function

Sub CopyFiles(SrcFolder, DstFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, UseDynamicStarNameSubfolder)
	Dim d, f
	Dim ShortCombinedDstFolder, CombinedDstFolder, folder_exists
	Dim iFN, oFN
	Dim ss
	
	If Not UseDynamicStarNameSubfolder Then
		CombinedDstFolder = DstFolder & DstFolderSuffix
		folder_exists = FolderExists(CombinedDstFolder)
	Else
		folder_exists = FolderExists(DstFolder)
	End If
	If folder_exists Then
		For Each f In oFSO.GetFolder(SrcFolder).Files
			If StrComp(oFSO.GetExtensionName(f), Extension) = 0 Then
				count = count + 1
				iFN = oFSO.GetFileName(f)
				ss = Len(SearchStr) = 0 Or InStr(iFN, SearchStr) > 0
				If ss Then
					If UseDynamicStarNameSubfolder Then 
						ShortCombinedDstFolder = DstFolder & StarName(iFN, Extension)
						CombinedDstFolder = ShortCombinedDstFolder & DstFolderSuffix
					End If
					oFN = DstFilenamePrefix & Replace(iFN, oldStr, newStr)
					If first Then first = FirstOperation(False, SrcFolder & iFN, CombinedDstFolder & oFN, False)
					If UseDynamicStarNameSubfolder Then 
						If Not oFSO.FolderExists(ShortCombinedDstFolder) Then 
							logfile.WriteLine "create folder: " & ShortCombinedDstFolder
							oFSO.CreateFolder ShortCombinedDstFolder
						End If
						If Not oFSO.FolderExists(CombinedDstFolder) Then 
							logfile.WriteLine "create folder: " & CombinedDstFolder
							oFSO.CreateFolder CombinedDstFolder
						End If
					End If
					
					If Not oFSO.FileExists(CombinedDstFolder & oFN) Then
						logfile.WriteLine "copy " & SrcFolder & iFN & " to " & CombinedDstFolder & oFN
						oFSO.CopyFile SrcFolder & iFN, CombinedDstFolder & oFN, overWrite
						moved_copied = moved_copied + 1
					Else
						If overWrite Then
							logfile.WriteLine "copy(replace) " & SrcFolder & iFN & " to " & CombinedDstFolder & oFN
							oFSO.CopyFile SrcFolder & iFN, CombinedDstFolder & oFN, overWrite
							replaced = replaced + 1
						Else
							' files are not repalced, skiping
							alreadyExists = alreadyExists + 1
						End If
					End If
				Else
					skipped = skipped + 1
				End If
			End If
		Next
	End If
End Sub

Sub MoveFiles(SrcFolder, DstFolder, DstFolderSuffix, DstFilenamePrefix, Extension, SearchStr, oldStr, newStr, overWrite, UseDynamicStarNameSubfolder)
	Dim d, f
	Dim ShortCombinedDstFolder, CombinedDstFolder, folder_exists
	Dim iFN, oFN, i, o
	Dim ss
	Dim Rename, action
	
	Rename = (SrcFolder = DstFolder)
	If Rename Then
		action = "rename "
	Else
		action = "move "
	End If
	
	If Not UseDynamicStarNameSubfolder Then
		CombinedDstFolder = DstFolder & DstFolderSuffix
		folder_exists = FolderExists(CombinedDstFolder)
	Else
		folder_exists = FolderExists(DstFolder)
	End If
	
	If folder_exists Then
		For Each f In oFSO.GetFolder(SrcFolder).Files
			If StrComp(oFSO.GetExtensionName(f), Extension) = 0 Then
				count = count + 1
				iFN = oFSO.GetFileName(f)
				ss = Len(SearchStr) = 0 Or InStr(iFN, SearchStr) > 0
				If ss Then
					If UseDynamicStarNameSubfolder Then
						ShortCombinedDstFolder = DstFolder & StarName(iFN, Extension)
						CombinedDstFolder = ShortCombinedDstFolder & DstFolderSuffix
					End If
					oFN = DstFilenamePrefix & Replace(iFN, oldStr, newStr)
					i = SrcFolder & iFN
					o = CombinedDstFolder & oFN
					If first Then first = FirstOperation(True, i, o, Rename)
					If UseDynamicStarNameSubfolder Then 
						If Not oFSO.FolderExists(ShortCombinedDstFolder) Then 
							logfile.WriteLine "create folder: " & ShortCombinedDstFolder
							oFSO.CreateFolder ShortCombinedDstFolder
						End If
						If Not oFSO.FolderExists(CombinedDstFolder) Then 
							logfile.WriteLine "create folder: " & CombinedDstFolder
							oFSO.CreateFolder CombinedDstFolder
						End If
					End If
					
					If InStr(logFN, iFN) = 0 And InStr(logFN, oFN) = 0 Then
						If Not oFSO.FileExists(o) Then 
							logfile.WriteLine action & i & " to " & o
							oFSO.MoveFile i, o
							moved_copied = moved_copied + 1
						Else
							If overWrite Then
								If i <> o Then
									logfile.WriteLine action & "(replace) " & i & " to " & o
									oFSO.DeleteFile o
									oFSO.MoveFile i, o
									replaced = replaced + 1
								Else
									logfile.WriteLine "skip " & i & " to " & o
									alreadyExists = alreadyExists + 1
								End If
							Else
								' files are not repalced, skiping
								alreadyExists = alreadyExists + 1
							End If
						End If
					End If
				Else
					skipped = skipped + 1
				End If
			End If
		Next
	End If
End Sub

' delete-empty-subfolders.WSF
Sub DeleteEmptySubfolders(RootFolder)
	Dim d, s, f, i, count, str, DN
	
	RootFolder = FolderNameWithSlash(RootFolder)
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	str = "Empty subfolders of folder " & RootFolder & " will be deleted." & vbNewLine & vbNewLine & "Do you want to proceed?"
	If vbNo = MsgBox(str, vbYesNo, "Delete Subfolders") Then
		WScript.Echo "Script aborted by user. No folder deleted."
		WScript.Quit
	End If
	
	If oFSO.FolderExists(RootFolder) Then
		OpenLogFile RootFolder
		count = 0
		For Each d In oFSO.GetFolder(RootFolder).SubFolders
			i = 0
			For Each f In d.Files
				i = i + 1
			Next
			For Each s In d.SubFolders
				i = i + 1
			Next
			If i = 0 Then
				DN = RootFolder & d.Name
				oFSO.DeleteFolder DN
				count = count + 1
				logfile.WriteLine "delete: " & DN
			Else
				logfile.WriteLine "skip: " & DN
			End If
		Next
		logfile.Close
		Set logfile = Nothing
		MsgBox "All done:" & vbNewLine & vbNewLine & count & " empty subfolders deleted in folder " & RootFolder, vbOK, "Summary"
	Else
		WScript.Echo "Folder does not exist: " & RootFolder
	End If
End Sub
