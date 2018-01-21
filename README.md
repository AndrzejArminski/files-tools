
# files-tools
**Repository of VBS scripts to count, copy, move, delete and rename files**

## Scripts in this Repository:

files-tools.WSF
copy-files-folder-to-folder.WSF  
copy-files-subfolders-to-folder.WSF  
copy-files-folder-to-subfolders.WSF  
copy-files-subfolders-to-subfolders.WSF  
move-files-folder-to-folder.WSF  
move-files-subfolders-to-folder.WSF  
move-files-folder-to-subfolders.WSF  
move-files-subfolders-to-subfolders.WSF  
delete-files-from-folder.WSF  
delete-files-from-subfolders.WSF  
delete-empty-subfolders.WSF  
rename-files-in-folder.WSF  
rename-files-in-subfolders.WSF  
count-files-in-folder.WSF  
count-files-in-subfolders.WSF  

Supply script arguments by editing constants in .WSF files.  
Scripts create log file in destination folder with name: files-tools-logfile-yyyy-mm-dd-hh-mm-ss.txt.

___
### files-tools.WSF
You may use files-tools.WSF to perform all actions of counting, copying, moving, renaming and deleting files:

**copy|move|copyOverwrite|moveOverwrite F:\SrcFolder\<<AnyFolder>>\SrcFolderSuffix\SearchStr^OldStr.Ext F:\DstFolder\<<StarNameFolder>>\DstFolderSuffix\FNPrefix^NewStr**  
**rename F:\Folder\<<AnyFolder>>\FolderSuffix\SearchStr^OldStr^NewStr.Ext**  
**delete|count F:\Folder\<<AnyFolder>>\FolderSuffix\SearchStr.Ext**  

Common arguments are:  
>**SourceFolder** - Single source folder containing files to be copied.  
**SourceRootFolder** - Folder with subfolders containing files to be moved.  
**SourceFolderSuffix** - SourceRootFolder's any subfolder's subfolder name. By supplying non-empty SourceFolderSuffix you move files from SourceRootFolder's subfolder's subfolder: SourceRootFolder\AnySubfolder\SourceFolderSuffix\.  
**DestinationFolder** - Single destination folder.  
**DestinationFolderSuffix** - DestinationRootFolder's StarName subfolder's subfolder name. By supplying non-empty DestinationFolderSuffix you copy files to DestinationRootFolder's subfolder's subfolder: DestinationRootFolder\StarNameSubfolder\DestinationFolderSuffix\.  
**Extension** - Only files with this extension in the source foulders will be processed. Required. No wildcards.  
**SearchString** - Only files with SearchString within file name will be processed. Give empty string to process all files.  
**OldString**, **NewString** - Replace a non-empty OldString with NewString in destination file names.  
**DestinationFilenamePrefix** - String to be attached at the begining of destination file names.  
**Overwrite** - When *True*, files in destination will be overwritten.

___
### copy-files-folder-to-folder.WSF
**CopyFilesFolderToFolder (SourceFolder, DestinationFolder,**
**DestinationFilenamePrefix, Extension, SearchString, OldString, NewString, Overwrite)**

Copies files with Extension and SearchString in file names from SourceFolder to DestinationFolder.
>Arguments:  
**SourceFolder** - Single source folder containing files to be copied.  
**DestinationFolder** - Single destination folder.

___
### copy-files-subfolders-to-folder.WSF
**CopyFilesSubfoldersToFolder (SourceRootFolder, SourceFolderSuffix, DestinationFolder,**
**Extension, SearchString, OldString, NewString, DestinationFilenamePrefix, Overwrite)**

Copies files with Extension and SearchString in file names from SourceRootFolder's subfolders to DestinationFolder.
>Arguments:  
**SourceRootFolder** - Folder with subfolders containing files to be copied.  
**SourceFolderSuffix** - SourceRootFolder's any subfolder's subfolder name. By supplying non-empty SourceFolderSuffix you copy files from SourceRootFolder's subfolder's subfolder: (SourceRootFolder\AnySubfolder\SourceFolderSuffix\).  
**DestinationFolder** - Single destination folder.

___
### copy-files-folder-to-subfolders.WSF
**CopyFilesFolderToSubfolders (SourceFolder, DestinationRootFolder, DestinationFolderSuffix,**
**Extension, SearchString, OldString, NewString, DestinationFilenamePrefix, Overwrite)**

Copies files with Extension and SearchString in file names from SourceFolder to DestinationRootFolder's subfolders **using StarNames** as subfolder's names. StarName is derived from file name, as the last underscore delimited token.
>Arguments:  
**SourceFolder** - Single folder containing files to be copied.  
**DestinationRootFolder** - Destination root folder. Files will be copied to DestinationRootFolder's subfolders named after StarNames derived from file name.  
**DestinationFolderSuffix** - DestinationRootFolder's StarName subfolder's subfolder name. By supplying non-empty DestinationFolderSuffix you copy files to DestinationRootFolder's subfolder's subfolder: (DestinationRootFolder\StarNameSubfolder\DestinationFolderSuffix\).

___
### copy-files-subfolders-to-subfolders.WSF
**CopyFilesSubfoldersToSubfolders (SourceRootFolder, SourceFolderSuffix, DestinationRootFolder, DestinationFolderSuffix,**
**Extension, SearchString, OldString, NewString, DestinationFilenamePrefix, Overwrite)**

Copies files with Extension and SearchString in file names from SourceRootFolder's subfolders to DestinationRootFolder's subfolders **using StarNames** as destination subfolder's names. StarName is derived from file name, as the last underscore delimited token.
>Arguments:  
**SourceFolder** - Single folder containing files to be copied.  
**SourceFolderSuffix** - SourceRootFolder's any subfolder's subfolder name. By supplying non-empty SourceFolderSuffix you copy files from SourceRootFolder's subfolder's subfolder: (SourceRootFolder\AnySubfolder\SourceFolderSuffix\).  
**DestinationRootFolder** - Destination folder. Files will be copied to subfolders named after StarNames derived from file name.  
**DestinationFolderSuffix** - DestinationRootFolder's StarName subfolder's subfolder name. By supplying non-empty DestinationFolderSuffix you copy files to DestinationRootFolder's subfolder's subfolder: (DestinationRootFolder\StarNameSubfolder\DestinationFolderSuffix\).

___
### move-files-folder-to-folder.WSF
**MoveFilesFolderToFolder (SourceFolder, DestinationFolder,**
**Extension, SearchString, OldString, NewString, DestinationFilenamePrefix, Overwrite)**

Moves files with Extension and SearchString in file names from SourceFolder to DestinationFolder.
>Arguments:  
**SourceFolder** - Single source folder containing files to be moved.  
**DestinationFolder** - Single destination folder.

___
### move-files-subfolders-to-folder.WSF
**MoveFilesSubfoldersToFolder (SourceRootFolder, SourceFolderSuffix, DestinationFolder,**
**Extension, SearchString, OldString, NewString, DestinationFilenamePrefix, Overwrite)**

Moves files with Extension and SearchString in file names from SourceRootFolder's subfolders to DestinationFolder.
>Arguments:  
**SourceRootFolder** - Folder with subfolders containing files to be moved.  
**SourceFolderSuffix** - SourceRootFolder's any subfolder's subfolder name. By supplying non-empty SourceFolderSuffix you move files from SourceRootFolder's subfolder's subfolder: (SourceRootFolder\AnySubfolder\SourceFolderSuffix\).  
**DestinationFolder** - Single destination folder.

___
### move-files-folder-to-subfolders.WSF
**MoveFilesFolderToSubfolders (SourceFolder, DestinationRootFolder, DestinationFolderSuffix,**
**Extension, SearchString, OldString, NewString, DestinationFilenamePrefix, Overwrite)**

Moves files with Extension and SearchString in file names from SourceFolder to DestinationRootFolder's subfolders **using StarNames** as subfolder's names. StarName is derived from file name, as the last underscore delimited token.
>Arguments:  
**SourceFolder** - Single folder containing files to be moved.  
**DestinationRootFolder** - Destination folder. Files will be moved to subfolders named after StarNames derived from file name.  
**DestinationFolderSuffix** - DestinationRootFolder's StarName subfolder's subfolder name. By supplying non-empty DestinationFolderSuffix you move files to DestinationRootFolder's subfolder's subfolder: (DestinationRootFolder\StarNameSubfolder\DestinationFolderSuffix\).

___
### move-files-subfolders-to-subfolders.WSF
**MoveFilesSubfoldersToSubfolders (SourceRootFolder, SourceFolderSuffix, DestinationRootFolder, DestinationFolderSuffix,**
**Extension, SearchString, OldString, NewString, DestinationFilenamePrefix, Overwrite)**

Moves files with Extension and SearchString in file names from SourceRootFolder's subfolders to DestinationRootFolder's subfolders **using StarNames** as subfolder's names. StarName is derived from file name, as the last underscore delimited token.
>Arguments:  
**SourceRootFolder** - Folder with subfolders containing files to be moved.  
**SourceFolderSuffix** - SourceRootFolder's any subfolder's subfolder name. By supplying non-empty SourceFolderSuffix you move files from SourceRootFolder's subfolder's subfolder: (SourceRootFolder\AnySubfolder\SourceFolderSuffix\).  
**DestinationRootFolder** - Destination folder. Files will be moved to subfolders named after StarNames derived from file name.  
**DestinationFolderSuffix** - DestinationRootFolder's StarName subfolder's subfolder name. By supplying non-empty DestinationFolderSuffix you move files to DestinationRootFolder's subfolder's subfolder: (DestinationRootFolder\StarNameSubfolder\DestinationFolderSuffix\).

___
### delete-files-from-folder.WSF
**DeleteFilesFromFolder (Folder, Extension, SearchString)**

Deletes files with Extension and SearchString in file names from Folder.
>Arguments:  
**Folder** - Single folder containing files to be deleted.

___
### delete-files-from-subfolders.WSF
**DeleteFilesFromSubolders (RootFolder, Extension, SearchString)**

Deletes files with Extension and SearchString in file names from subfolders of RootFolder.
>Arguments:  
**RootFolder** - Folder with subfolders containing files to be deleted.

___
### delete-empty-subfolders.WSF
**DeleteEmptySubolders (RootFolder)**

Deletes subfolders of RootFolder that do not contain files.
>Arguments:  
**RootFolder** - Folder with empty subfolders to be deleted.

___
### rename-files-in-folder.WSF
**RenameFilesFromFolder (Folder, Extension, SearchString)**

Renames files with Extension and SearchString in file names, in Folder.
>Arguments:  
**Folder** - Single folder containing files to be renamed.

___
### rename-files-in-subfolders.WSF
**RenameFilesFromSubfolders (RootFolder, Extension, SearchString)**

Renames files with Extension and SearchString in file names, in subfolders of RootFolder.
>Arguments:  
**RootFolder** - Folder with subfolders containing files to be renamed.

___
### count-files-in-folder.WSF
**CountFilesInFolder (Folder, Extension, SearchString)**

Returns number of files with Extension and SearchString in file names, in Folder.
>Arguments:  
**Folder** - Single folder containing files to be counted.

___
### count-files-in-subfolders.WSF
**CountFilesInSubolders (RootFolder, Extension, SearchString)**

Returns number of files with Extension and SearchString in file names, in subfolders of RootFolder.
>Arguments:  
**RootFolder** - Folder with subfolders containing files to be counted.

___
### star-name
**Function StarName(FileName, Extension)**

Returns StarName string: last undescore delimited token in FileName. Result is used to name DestinationRootFolder's subfolder. Function StarName resides in  files-tools-LIB.vbs
>Function StarName(FileName, Extension)  
' For FileName: "SFDB_2016-02-03_1917-29_J000703_BMAH__V_1x1_0040s_HIP 110893.fit"  
' returns star name: "HIP 110893"  
  Dim t  
  t = Split(Left(FN, Len(FN) - (Len(Extension) + 1)), "_")  
  StarName = t(Ubound(t)) & "\"  
End Function

___
### files-tools-LIB.vbs
The file **files-tools-LIB.vbs** is a library containing subroutines used by *.WSF scripts.
