'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------
Option Explicit
'################################################
'This script is to convert office documents to PDF files
'################################################
Sub main()
    Dim ArgCount
    ArgCount = WScript.Arguments.Count
    Select Case ArgCount
    Case 1
        MsgBox "Please ensure documents are saved,if that press 'OK' to continue",,"Warning"
        Dim OfficeFilePaths,objshell
        OfficeFilePaths = WScript.Arguments(0)
        StopPptApp
        StopWordApp
        StopXlsApp
        Set objshell = CreateObject("scripting.filesystemobject")
        If objshell.FolderExists(OfficeFilePaths) Then
            Dim flag, FileNumber, OfficeFilePath, Folder, OfficeFiles, OfficeFile
            flag = 0
            FileNumber = 0
            Set Folder = objshell.GetFolder(OfficeFilePaths)
            Set OfficeFiles = Folder.Files
            For Each OfficeFile In OfficeFiles
                FileNumber=FileNumber+1
                OfficeFilePath = OfficeFile.Path
                If GetPptFile(OfficeFilePath) Then
                    ConvertPptToPDF OfficeFilePath
                    flag=flag+1
                Elseif GetWordFile(OfficeFilePath) Then
                    ConvertWordToPDF OfficeFilePath
                    flag=flag+1
                Elseif GetXlsFile(OfficeFilePath) Then
                    ConvertXlsToPDF OfficeFilePath
                    flag=flag+1
                End If
            Next
            WScript.Echo "Total " & FileNumber & " file(s) in the folder and convert " & flag & " Documents to PDF fles."

        Else
            If GetPptFile(OfficeFilePaths) Then
                ConvertPptToPDF OfficeFilePaths
            Elseif GetWordFile(OfficeFilePaths) Then
                ConvertWordToPDF OfficeFilePaths
            Elseif GetXlsFile(OfficeFilePaths) Then
                ConvertXlsToPDF OfficeFilePaths
            Else
                WScript.Echo "Please drag a document or a folder with office documents."
            End If
            WScript.Echo "Done."
        End If

    Case  Else
            WScript.Echo "Please drag a document or a folder with office documents."
    End Select
End Sub

Function ConvertWordToPDF(DocPath)
    Dim objshell,ParentFolder,BaseName,wordapp,doc,PDFPath
    Set objshell= CreateObject("scripting.filesystemobject")
    ParentFolder = objshell.GetParentFolderName(DocPath)
    BaseName = objshell.GetBaseName(DocPath)
    PDFPath = parentFolder & "\" & BaseName & ".pdf"
    Set wordapp = CreateObject("Word.application")
    Set doc = wordapp.documents.open(DocPath)
    doc.saveas PDFPath,17
    doc.close
    wordapp.quit
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
