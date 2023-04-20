Option Explicit
Dim objShell, objFSO, strUser, strSourceFolder, strDestinationFolder

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strUser = LCase(objShell.ExpandEnvironmentStrings("%USERNAME%"))

' Define the folders to copy
Dim folders(3)
folders(0) = "Documents"
folders(1) = "Videos"
folders(2) = "Music"
folders(3) = "Favorites"

Dim folder
For Each folder In folders
    strSourceFolder = "C:\Source\CP\" & strUser & "\" & folder
    strDestinationFolder = objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\" & folder

    If objFSO.FolderExists(strSourceFolder) Then
        CopyFolderContent strSourceFolder, strDestinationFolder
    Else
        WScript.Echo "Source folder not found: " & strSourceFolder
    End If
Next

Sub CopyFolderContent(srcFolder, dstFolder)
    Dim srcFile, dstFile
    For Each srcFile In objFSO.GetFolder(srcFolder).Files
        dstFile = dstFolder & "\" & objFSO.GetFileName(srcFile)
        objFSO.CopyFile srcFile, dstFile, True
    Next

    Dim srcSubFolder, dstSubFolder
    For Each srcSubFolder In objFSO.GetFolder(srcFolder).SubFolders
        dstSubFolder = dstFolder & "\" & objFSO.GetFileName(srcSubFolder)
        If Not objFSO.FolderExists(dstSubFolder) Then
            objFSO.CreateFolder dstSubFolder
        End If
        CopyFolderContent srcSubFolder, dstSubFolder
    Next
End Sub
