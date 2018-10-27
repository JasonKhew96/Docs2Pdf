Option Explicit

Sub main()
    Dim ArgCount
    ArgCount = WScript.Arguments.Count
    Select Case ArgCount
    Case 1
        Convert
    Case  Else
            WScript.Echo "Please drag a document or a folder with office documents."
    End Select
End Sub

Function Convert
    ' MsgBox "Please ensure documents are saved,if that press 'OK' to continue",,"Warning"
    Dim OfficeFilePaths, FileNumber
    OfficeFilePaths = WScript.Arguments(0)

    StopPptApp
    StopWordApp
    StopXlsApp

    FileNumber = ProcessFile(OfficeFilePaths)

    WScript.Echo "Total " & FileNumber & " file(s) converted to PDF."
End Function

Function ProcessFile(path)
    Dim fileNumber, objshell, folder, file, OfficeFiles, subFolders, fol
    fileNumber = 0
    Set objshell = CreateObject("scripting.filesystemobject")
    If objshell.FolderExists(path) Then
        Set folder = objshell.GetFolder(path)
        Set OfficeFiles = folder.Files
        For Each file In OfficeFiles
            if ConvertOneFile(file.path) Then
                fileNumber = fileNumber + 1
            End If
        Next
        Set subFolders = folder.SubFolders
        For Each fol In subFolders
            fileNumber = fileNumber + ProcessFile(fol)
        Next
    Elseif ConvertOneFile(file) Then
        fileNumber = fileNumber + 1
    End If
    ProcessFile = fileNumber
End Function

Function ConvertOneFile(path)
    If GetPptFile(path) Then
        ConvertPptToPDF path
        ConvertOneFile = true
    Elseif GetWordFile(path) Then
        ConvertWordToPDF path
        ConvertOneFile = true
    Elseif GetXlsFile(path) Then
        ConvertXlsToPDF path
        ConvertOneFile = true
    Else
        ConvertOneFile = false
    End If
End Function


Function ConvertWordToPDF(DocPath)
    Dim objshell,ParentFolder,BaseName,wordapp,doc,PDFPath
    Set objshell= CreateObject("scripting.filesystemobject")
    ParentFolder = objshell.GetParentFolderName(DocPath)
    BaseName = objshell.GetBaseName(DocPath)
    PDFPath = parentFolder & "\" & BaseName & ".pdf"
    Set wordapp = CreateObject("Word.application")
    wordapp.WordBasic.DisableAutoMacros
    Set doc = wordapp.Documents.Open(DocPath)
    doc.saveas PDFPath, 17
    doc.close
    wordapp.quit
    Set wordapp = Nothing
    Set objshell = Nothing
End Function

Function ConvertPptToPDF(PptPath)
    Dim objshell,ParentFolder,BaseName,ppapp,doc,PDFPath
    Set objshell= CreateObject("scripting.filesystemobject")
    ParentFolder = objshell.GetParentFolderName(PptPath)
    BaseName = objshell.GetBaseName(PptPath)
    PDFPath = parentFolder & "\" & BaseName & ".pdf"
    Set ppapp = CreateObject("PowerPoint.application")
    Set doc = ppapp.Presentations.open(PptPath, , , 0)
    doc.saveas PDFPath,32
    doc.close
    ppapp.quit
    Set objshell = Nothing
End Function

Function ConvertXlsToPDF(XlsPath)
    Dim objshell,ParentFolder,BaseName,xlsapp,doc,PDFPath
    Set objshell= CreateObject("scripting.filesystemobject")
    ParentFolder = objshell.GetParentFolderName(XlsPath)
    BaseName = objshell.GetBaseName(XlsPath)
    PDFPath = parentFolder & "\" & BaseName & ".pdf"
    Set xlsapp = CreateObject("Excel.application")
    Set doc = xlsapp.Workbooks.Open(XlsPath)
    doc.ExportAsFixedFormat 0, PDFPath
    doc.saved = True
    doc.close
    xlsapp.quit
    Set objshell = Nothing
End Function

Function GetWordFile(DocPath)
    Dim objshell
    Set objshell= CreateObject("scripting.filesystemobject")
    Dim Arrs ,Arr
    Arrs = Array("doc","docx")
    Dim blnIsDocFile,FileExtension
    blnIsDocFile= False
    FileExtension = objshell.GetExtensionName(DocPath)
    For Each Arr In Arrs
        If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then
            blnIsDocFile= True
            Exit For
        End If
    Next
    GetWordFile = blnIsDocFile
    Set objshell = Nothing
End Function

Function GetPptFile(PptPath)
    Dim objshell
    Set objshell= CreateObject("scripting.filesystemobject")
    Dim Arrs ,Arr
    Arrs = Array("ppt","pptx")
    Dim blnIsPptFile,FileExtension
    blnIsPptFile= False
    FileExtension = objshell.GetExtensionName(PptPath)
    For Each Arr In Arrs
        If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then
            blnIsPptFile= True
            Exit For
        End If
    Next
    GetPptFile = blnIsPptFile
    Set objshell = Nothing
End Function

Function GetXlsFile(XlsPath)
    Dim objshell
    Set objshell= CreateObject("scripting.filesystemobject")
    Dim Arrs ,Arr
    Arrs = Array("xls","xlsx")
    Dim blnIsPxlsFile,FileExtension
    blnIsPxlsFile= False
    FileExtension = objshell.GetExtensionName(XlsPath)
    For Each Arr In Arrs
        If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then
            blnIsPxlsFile= True
            Exit For
        End If
    Next
    GetXlsFile = blnIsPxlsFile
    Set objshell = Nothing
End Function

Function StopWordApp
    Dim strComputer,objWMIService,colProcessList,objProcess
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name = 'Winword.exe'")
    For Each objProcess in colProcessList
        objProcess.Terminate()
    Next
End Function

Function StopPptApp
    Dim strComputer,objWMIService,colProcessList,objProcess
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name = 'PowerPnt.exe'")
    For Each objProcess in colProcessList
        objProcess.Terminate()
    Next
End Function

Function StopXlsApp
    Dim strComputer,objWMIService,colProcessList,objProcess
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name = 'Excel.exe'")
    For Each objProcess in colProcessList
        objProcess.Terminate()
    Next
End Function

Call main
