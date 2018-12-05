Option Explicit

Sub main()
    Dim ArgCount
    ArgCount = WScript.Arguments.Count
    Select Case ArgCount
    Case 1
        Call convert
    Case Else
        WScript.Echo "Please drag a document or a folder to this vbs script."
    End Select
End Sub

Function FileExists(FilePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
        FileExists=CBool(1)
    Else
        FileExists=CBool(0)
    End If
End Function

Function convert
    MsgBox "Please ensure documents are saved, if that press 'OK' to continue",,"Warning"
    Dim Args1, TotalFiles
    Args1 = WScript.Arguments(0)

    TotalFiles = ProcessFile(Args1)

    WScript.Echo "Total " & TotalFiles & " file(s) converted to PDF."
End Function

Function ProcessFile(Path)
    Dim TotalFiles, Objshell
    TotalFiles = 0
    Set Objshell = CreateObject("scripting.filesystemobject")
    If Objshell.FolderExists(Path) Then
    Dim FolderPath, File, Files, SubFolder, Folder
        Set FolderPath = Objshell.GetFolder(Path)
        Set Files = FolderPath.Files
        For Each File In Files
            if convertOfficeFile(File.Path) Then
                TotalFiles = TotalFiles + 1
            End If
        Next
        Set SubFolder = FolderPath.SubFolders
        For Each Folder In SubFolder
            TotalFiles = TotalFiles + ProcessFile(Folder)
        Next
    Elseif convertOfficeFile(Path) Then
        TotalFiles = TotalFiles + 1
    End If
    ProcessFile = TotalFiles
End Function

Function convertOfficeFile(Path)
    If isPptFile(Path) Then
        convertPptToPDF Path
        convertOfficeFile = true
    Elseif isWordFile(Path) Then
        convertWordToPDF Path
        convertOfficeFile = true
    Elseif isXlsFile(Path) Then
        convertXlsToPDF Path
        convertOfficeFile = true
    Else
        convertOfficeFile = false
    End If
End Function


Function convertWordToPDF(Path)
    Dim Objshell, ParentFolder, BaseName, WordApp, Doc, PDFPath
    Set Objshell = CreateObject("scripting.filesystemobject")
    ParentFolder = Objshell.GetParentFolderName(Path)
    BaseName = Objshell.GetBaseName(Path)
    PDFPath = parentFolder & "\" & BaseName & ".pdf"
    If not FileExists(PDFPath) Then
        Set WordApp = CreateObject("Word.application")
        WordApp.WordBasic.DisableAutoMacros
        Set Doc = WordApp.Documents.Open(Path)
        Doc.saveas PDFPath, 17
        Doc.close
        WordApp.quit
    End If
    Set Objshell = Nothing
End Function

Function convertPptToPDF(Path)
    Dim Objshell, ParentFolder, BaseName, PptApp, Doc, PDFPath
    Set Objshell = CreateObject("scripting.filesystemobject")
    ParentFolder = Objshell.GetParentFolderName(Path)
    BaseName = Objshell.GetBaseName(Path)
    PDFPath = parentFolder & "\" & BaseName & ".pdf"
    If not FileExists(PDFPath) Then
        Set PptApp = CreateObject("PowerPoint.application")
        Set Doc = PptApp.Presentations.open(Path, , , 0)
        Doc.saveas PDFPath,32
        Doc.close
        PptApp.quit
    End If
    Set Objshell = Nothing
End Function

Function convertXlsToPDF(Path)
    Dim Objshell, ParentFolder, BaseName, XlsApp, Doc, PDFPath
    Set Objshell = CreateObject("scripting.filesystemobject")
    ParentFolder = Objshell.GetParentFolderName(Path)
    BaseName = Objshell.GetBaseName(Path)
    PDFPath = parentFolder & "\" & BaseName & ".pdf"
    If not FileExists(PDFPath) Then
        Set XlsApp = CreateObject("Excel.application")
        Set Doc = XlsApp.Workbooks.Open(Path)
        Doc.ExportAsFixedFormat 0, PDFPath
        Doc.saved = True
        Doc.close
        XlsApp.quit
    End If
    Set Objshell = Nothing
End Function

Function isWordFile(Path)
    Dim Objshell
    Set Objshell = CreateObject("scripting.filesystemobject")
    Dim Arrs, Arr
    Arrs = Array("doc","docx")
    Dim FileExtension
    isWordFile = False
    FileExtension = Objshell.GetExtensionName(Path)
    For Each Arr In Arrs
        If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then
            isWordFile = True
            Exit For
        End If
    Next
    Set Objshell = Nothing
End Function

Function isPptFile(Path)
    Dim Objshell
    Set Objshell = CreateObject("scripting.filesystemobject")
    Dim Arrs, Arr
    Arrs = Array("ppt","pptx")
    Dim FileExtension
    isPptFile = False
    FileExtension = Objshell.GetExtensionName(Path)
    For Each Arr In Arrs
        If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then
            isPptFile = True
            Exit For
        End If
    Next
    Set Objshell = Nothing
End Function

Function isXlsFile(Path)
    Dim Objshell
    Set Objshell = CreateObject("scripting.filesystemobject")
    Dim Arrs, Arr
    Arrs = Array("xls","xlsx")
    Dim FileExtension
    isXlsFile = False
    FileExtension = Objshell.GetExtensionName(Path)
    For Each Arr In Arrs
        If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then
            isXlsFile = True
            Exit For
        End If
    Next
    Set Objshell = Nothing
End Function

Call main
