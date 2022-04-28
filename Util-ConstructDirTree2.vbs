Option Explicit

Rem Command
Rem >cscript Util-ConstructDirTree2.vbs "C:\_seus\_work\t e m p" \a b\d\e \e\f\g\ "\h i\j k" lmn\opq
Rem  ------- -------------------------- ------------------------ -- ----- ------- ---------- -------
Rem                                     First(0)                 (1)  (2)    (3)     (4)      (5)

Rem Pass the first argument as the root folder where sub folders will be created to.
Rem Set the second argument and forth as relative folder path.
Rem Set relative path with or without preceding backslash
Rem Do not use spaces in folder name. If you must, close with double quotes.
Rem Trailing backslash does not cause problem. You may use with or without trailing backslash.
Rem 
Rem Result tree
Rem 
Rem +---a
Rem +---b
Rem |   \---d
Rem |       \---e
Rem +---e
Rem |   \---f
Rem |       \---g
Rem +---h i
Rem |   \---j k
Rem \---lmn
Rem     \---opq

Rem - File System Object
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")

Rem - Root Folder as the first argument
Dim strDirRoot: strDirRoot = WScript.Arguments(0)

Rem - Ensure the passed root folder exists in the system.
Rem   If root folder exits, delete the folder first for cleaning and then recreate it.
Rem   If root folder does not exist, quit.
If objFSO.FolderExists(strDirRoot) Then
	WScript.Echo "Input folder is vaid."
	Rem Microsoft VBScript runtime error: Permission denied
	Rem ===================================================
	Rem If a file is open in the root directory, following error will go off. 
	Rem Ensure all files are closed.
	Rem Also, if the directory is opened with command prompt, this error will go off. 
	objFSO.DeleteFolder(strDirRoot)
	WScript.Echo "strDirRoot was deleted for cleaning."
	objFSO.CreateFolder(strDirRoot)
	WScript.Echo "strDirRoot was recreated."
Else
	Rem If provided folder does not exist, script should stop there.
	WScript.Echo "ERROR Folder Not Exist. strDirRoot=" & strDirRoot
	WScript.Quit 1
End If

Rem Recursive call to create a nest folder structure.
Sub CreateSubFolder(ByVal argAbsPath)
	If Not objFSO.FolderExists(argAbsPath) Then
		CreateSubFolder objFSO.GetParentFolderName(argAbsPath)
		objFSO.CreateFolder argAbsPath
		WScript.Echo "Created Folder: argAbsPath=" & argAbsPath
	End If
End Sub

Rem - Loop through passed arguments, which are relative folder paths preceded with backslash
WScript.Echo "WScript.Arguments.Count=" & WScript.Arguments.Count
If WScript.Arguments.Count > 1 Then
	Dim i
	Dim strAbsPath
	Dim strRelPath
	For i = 1 To WScript.Arguments.Count - 1
		Rem - if relative path does not start with backslash, prepend a backslash.
		strRelPath = WScript.Arguments(i)
		If Not Left(strRelPath, 1) = "\" Then strRelPath = "\" & strRelPath
		Rem - construct an absolute folder path
		strAbsPath = strDirRoot & strRelPath
		WScript.Echo "strAbsPath=" & strAbsPath
		Rem - create the absolute folder path
		CreateSubFolder strAbsPath
	Next
End If

Set objFSO = Nothing


