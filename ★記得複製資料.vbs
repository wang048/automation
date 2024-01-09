Option Explicit
Dim fso, sourceDir1, sourceDir2, sourceDir3, targetDirBase, targetDir, newFolderName
Dim newestFile1, newestFile2, newestFile3
Dim file, latestDate1, latestDate2, targetPath

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Define the source folders and the base target folder
sourceDir1 = "\\lei_svr\Adm\�H�O�귽��\31 ��ڤH��\02.�l���q LUS\02.��´��"
sourceDir2 = "\\lei_svr\Adm\�H�O�귽��\31 ��ڤH��\03.�l���q LCA\02.��´��"
sourceDir3 = "\\lei_svr\Adm\�H�O�귽��\31 ��ڤH��\01.�l���q LS\06.��´��"
targetDirBase = "\\lei_svr\Adm\�H�O�귽��\05 ��´�޲z\01 ��´��\2024 HQ��´��\���~�l���q��´��\"

' Prompt for new folder name
newFolderName = InputBox("�п�J�s��Ƨ����W��:", "�s��Ƨ��W��")
If newFolderName = "" Then
    WScript.Echo "����J��Ƨ��W�١A�ާ@�����C"
    WScript.Quit
End If

' Construct the full target directory path
targetDir = targetDirBase & newFolderName

' Create the new folder if it does not exist
If Not fso.FolderExists(targetDir) Then
    fso.CreateFolder(targetDir)
End If

' Function to get the newest PowerPoint file in a directory
Function GetNewestPowerPointFile(folder)
    '... [Function body remains the same]
End Function

' Function to get the newest PowerPoint file in a directory
Function GetNewestPowerPointFile(folder)
    Dim newestFile, latestDate, file
    Set newestFile = Nothing
    latestDate = DateSerial(1900, 1, 1)
    
    For Each file in fso.GetFolder(folder).Files
        If (LCase(fso.GetExtensionName(file.Name)) = "pptx" Or LCase(fso.GetExtensionName(file.Name)) = "ppt") And file.DateLastModified > latestDate Then
            Set newestFile = file
            latestDate = file.DateLastModified
        End If
    Next
    
    Set GetNewestPowerPointFile = newestFile
End Function

' Copy the newest PowerPoint file from the first source directory
Set newestFile1 = GetNewestPowerPointFile(sourceDir1)
If Not newestFile1 Is Nothing Then
    targetPath = fso.BuildPath(targetDir, newestFile1.Name)
    newestFile1.Copy targetPath, True
    WScript.Echo "Copied " & newestFile1.Name & " to " & targetDir
Else
    WScript.Echo "No PowerPoint files found in " & sourceDir1
End If

' Copy the newest PowerPoint file from the second source directory
Set newestFile2 = GetNewestPowerPointFile(sourceDir2)
If Not newestFile2 Is Nothing Then
    targetPath = fso.BuildPath(targetDir, newestFile2.Name)
    newestFile2.Copy targetPath, True
    WScript.Echo "Copied " & newestFile2.Name & " to " & targetDir
Else
    WScript.Echo "No PowerPoint files found in " & sourceDir2
End If

' Copy the newest PowerPoint file from the 3rd source directory
Set newestFile3 = GetNewestPowerPointFile(sourceDir3)
If Not newestFile3 Is Nothing Then
    targetPath = fso.BuildPath(targetDir, newestFile3.Name)
    newestFile3.Copy targetPath, True
    WScript.Echo "Copied " & newestFile3.Name & " to " & targetDir
Else
    WScript.Echo "No PowerPoint files found in " & sourceDir3
End If

