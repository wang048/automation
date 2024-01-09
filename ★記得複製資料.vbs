Option Explicit
Dim fso, sourceDir1, sourceDir2, sourceDir3, targetDirBase, targetDir, newFolderName
Dim newestFile1, newestFile2, newestFile3
Dim file, latestDate1, latestDate2, targetPath

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Define the source folders and the base target folder
sourceDir1 = "\\lei_svr\Adm\人力資源部\31 國際人管\02.子公司 LUS\02.組織圖"
sourceDir2 = "\\lei_svr\Adm\人力資源部\31 國際人管\03.子公司 LCA\02.組織圖"
sourceDir3 = "\\lei_svr\Adm\人力資源部\31 國際人管\01.子公司 LS\06.組織圖"
targetDirBase = "\\lei_svr\Adm\人力資源部\05 組織管理\01 組織圖\2024 HQ組織圖\海外子公司組織圖\"

' Prompt for new folder name
newFolderName = InputBox("請輸入新資料夾的名稱:", "新資料夾名稱")
If newFolderName = "" Then
    WScript.Echo "未輸入資料夾名稱，操作取消。"
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

