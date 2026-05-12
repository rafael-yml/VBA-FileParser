VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID) As Long
#End If

Private Const MaxRecursionDepth As Long = 5
Private Const MaxFileSizeBytes As Long = 30& * 1024 * 1024

Private m_PdfParser As PdfParser
Private m_WinOCR As WinOCR
Private m_WdCOM As WdCOM
Private m_Outlook As Object
Private m_CurrentType As String
Private m_CurrentDepth As Long
Private m_ExtractAttachments As Boolean
Private m_LastStatus As Long
Private m_TempFolder As String

Public Property Get Parser() As PdfParser
    If m_PdfParser Is Nothing Then
        Set m_PdfParser = New PdfParser
    End If
    Set Parser = m_PdfParser
End Property

Public Property Get LastType() As String
    LastType = m_CurrentType
End Property

Public Property Get LastStatus() As Long
    LastStatus = m_LastStatus
End Property

Private Sub Class_Initialize()
    m_CurrentType = ""
    m_CurrentDepth = 0
    m_ExtractAttachments = True
    m_LastStatus = 0
    m_TempFolder = ""
End Sub

Public Function ExtractText(ByVal filePath As String, Optional ByVal extractAttachments As Boolean = True) As String
    Dim pages As Collection
    Set pages = ExtractPages(filePath, extractAttachments)
    
    If pages.count > 0 Then
        Dim i As Long
        For i = 1 To pages.count
            If Len(pages(i)) > 0 Then
                ExtractText = ExtractText & pages(i) & vbCrLf
            End If
        Next
    End If
    
    If Len(ExtractText) = 0 And m_LastStatus = 0 Then
        m_LastStatus = 99
    End If
End Function

Public Function ExtractPages(ByVal filePath As String, Optional ByVal extractAttachments As Boolean = True) As Collection
    Dim extracted As Collection
    m_LastStatus = 0
    m_CurrentDepth = 0
    m_ExtractAttachments = extractAttachments
    m_CurrentType = ""
    m_TempFolder = ""
    Set extracted = New Collection

    If Len(Dir(filePath)) = 0 Then
        m_CurrentType = "Unknown"
        m_LastStatus = 1
        Exit Function
    End If
    
    Dim fSize As Long
    On Error Resume Next
    fSize = FileLen(filePath)
    On Error GoTo 0
    If fSize > MaxFileSizeBytes Then
        m_CurrentType = "TooLarge"
        m_LastStatus = 2
        Exit Function
    End If
    
    Dim ext As String
    ext = LCase(GetExtension(filePath))
    
    Select Case ext
        Case "pdf"
            Set extracted = ExtractPdfPages(filePath)
        Case "xls", "xlsx", "xlsb", "xlsm"
            Set extracted = ExtractExcelPages(filePath)
        Case "doc", "docx", "docm"
            Set extracted = ExtractWordPages(filePath)
        Case "txt", "csv"
            Set extracted = ExtractTextPages(filePath)
        Case "html", "htm"
            Set extracted = ExtractHtmlPages(filePath)
        Case "xml"
            Set extracted = ExtractXmlPages(filePath)
        Case "png", "jpg", "jpeg", "gif", "bmp", "webp", "tiff", "tif"
            Set extracted = ExtractImagePages(filePath)
        Case "msg"
            Set extracted = ExtractMsgPages(filePath)
        Case "rtf"
            Set extracted = ExtractRtfPages(filePath)
        Case "eml"
            Set extracted = ExtractEmlPages(filePath)
        Case "zip"
            Set extracted = ExtractZipPages(filePath)
    End Select
    
    If extracted.count = 0 And m_LastStatus = 0 Then
        m_LastStatus = 99
    End If
    
    Set ExtractPages = extracted
    CleanupTempFolder
End Function

Private Function GetExtension(ByVal filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, ".")
    If pos > 0 Then
        GetExtension = Mid(filePath, pos + 1)
    Else
        GetExtension = ""
    End If
End Function

Private Function GetTempFolder() As String
    If Len(m_TempFolder) = 0 Then
        m_TempFolder = Environ("TEMP") & "\fp_" & Format(Now, "YYYYMMDDHHMMSS") & Right$(Format(Timer, "0000"), 4)
    End If
    EnsureFolder m_TempFolder
    GetTempFolder = m_TempFolder
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Sub CleanupTempFolder()
    If Len(m_TempFolder) > 0 And Len(Dir(m_TempFolder, vbDirectory)) > 0 Then
        On Error Resume Next
        Kill m_TempFolder & "\*.*"
        RmDir m_TempFolder
        On Error GoTo 0
        m_TempFolder = ""
    End If
End Sub

Private Function NewGuid() As String
    Dim g As GUID
    Dim s As String
    CoCreateGuid g
    s = Right$("00000000" & Hex$(g.Data1), 8) & Right$("0000" & Hex$(g.Data2), 4) & Right$("0000" & Hex$(g.Data3), 4)
    NewGuid = s
End Function

Private Function GetOutlookInstance() As Object
    If Not m_Outlook Is Nothing Then
        Set GetOutlookInstance = m_Outlook
        Exit Function
    End If
    
    On Error Resume Next
    Set m_Outlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    If m_Outlook Is Nothing Then
        On Error Resume Next
        Set m_Outlook = CreateObject("Outlook.Application")
        On Error GoTo 0
    End If
    
    Set GetOutlookInstance = m_Outlook
End Function

Private Function ExtractPdfPages(ByVal filePath As String) As Collection
    Dim pdfPages As Collection
    Set pdfPages = New Collection
    m_CurrentType = "Pdf"
    Set pdfPages = Parser.ExtractPages(filePath)
    m_LastStatus = Parser.LastStatus
    Set ExtractPdfPages = pdfPages
End Function

Private Function ExtractExcelPages(ByVal filePath As String) As Collection
    Dim excelPages As Collection
    m_CurrentType = "Excel"
    m_LastStatus = 0
    Set excelPages = New Collection
    Dim wb As Object
    Dim ws As Object
    Dim sheetIndex As Long
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim cellValue As Variant
    Dim sheetText As String
    Dim emptyCells As Long
    Dim emptyRows As Long
    
    On Error Resume Next
    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    On Error GoTo 0
    
    If wb Is Nothing Then
        m_LastStatus = 10
        Set excelPages = FallbackToPdf(filePath)
        Exit Function
    End If
    
    For sheetIndex = 1 To wb.Worksheets.count
        Set ws = wb.Worksheets(sheetIndex)
        sheetText = ""
        
        On Error Resume Next
        lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).row
        lastCol = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
        On Error GoTo 0
        
        If lastRow = 0 Then lastRow = 1
        If lastCol = 0 Then lastCol = 1
        emptyRows = 0
        For r = 1 To lastRow
            emptyCells = 0
            For c = 1 To lastCol
                On Error Resume Next
                cellValue = ws.Cells(r, c).value
                On Error GoTo 0
                If Not IsEmpty(cellValue) Then
                    sheetText = sheetText & CStr(cellValue) & " "
                    emptyRows = 0
                Else
                    emptyCells = emptyCells + 1
                End If
            Next c
            If emptyCells >= lastCol - 1 Then emptyRows = emptyRows + 1
            If emptyRows >= 10 Then Exit For
            If lastCol > 0 And emptyRows = 0 Then sheetText = sheetText & vbCrLf
        Next r
        
        If Len(sheetText) > 0 Then
            excelPages.Add sheetText
        End If
    Next sheetIndex
    
    wb.Close SaveChanges:=False
    Set wb = Nothing
    
    If excelPages.count = 0 And m_CurrentDepth < MaxRecursionDepth Then
        Set excelPages = FallbackToPdf(filePath)
    End If
    
    If excelPages.count = 0 Then m_LastStatus = 11
    
    Set ExtractExcelPages = excelPages
    
End Function

Private Function ExtractWordPages(ByVal filePath As String) As Collection
    Dim wordPages As Collection
    m_CurrentType = "Word"
    m_LastStatus = 0
    Set wordPages = New Collection
    If m_WdCOM Is Nothing Then Set m_WdCOM = New WdCOM
    
    Dim text As String
    text = m_WdCOM.ExtractText(filePath)
    m_LastStatus = m_WdCOM.LastStatus
    
    If Len(text) > 0 Then
        wordPages.Add text
    End If
    
    If wordPages.count = 0 And m_CurrentDepth < MaxRecursionDepth Then
        Set wordPages = FallbackToPdf(filePath)
    End If
    
    If wordPages.count = 0 Then m_LastStatus = 21
    Set ExtractWordPages = wordPages
End Function

Private Function FallbackToPdf(ByVal filePath As String) As Collection
    Dim pdfPages As Collection
    m_CurrentDepth = m_CurrentDepth + 1
    If m_CurrentDepth > MaxRecursionDepth Then
        m_CurrentDepth = m_CurrentDepth - 1
        Exit Function
    End If
    Set pdfPages = New Collection
    
    Dim tempPdf As String
    tempPdf = GetTempFolder() & "_fallback.pdf"
    
    Dim oWord As Object
    Dim oDoc As Object
    
    On Error Resume Next
    Set oWord = GetObject(, "Word.Application")
    On Error GoTo 0
    
    If oWord Is Nothing Then
        On Error Resume Next
        Set oWord = CreateObject("Word.Application")
        On Error GoTo 0
    End If
    
    If oWord Is Nothing Then Exit Function
    
    oWord.Visible = False
    oWord.DisplayAlerts = 0
    
    On Error Resume Next
    Set oDoc = oWord.Documents.Open(filePath, ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False, NoEncodingDialog:=True)
    On Error GoTo 0
    
    If oDoc Is Nothing Then
        oWord.Quit
        Set oWord = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    oDoc.ExportAsFixedFormat OutputFileName:=tempPdf, ExportFormat:=17, OpenAfterExport:=False
    On Error GoTo 0
    
    oDoc.Close SaveChanges:=False
    oWord.Quit
    Set oDoc = Nothing
    Set oWord = Nothing
    
    If Len(Dir(tempPdf)) > 0 Then
        Set pdfPages = Parser.ExtractPages(tempPdf)
        m_CurrentType = "Pdf"
        On Error Resume Next
        Kill tempPdf
        On Error GoTo 0
    End If
    
    m_CurrentDepth = m_CurrentDepth - 1
    Set FallbackToPdf = pdfPages
End Function

Private Function ExtractTextPages(ByVal filePath As String) As Collection
    Dim txtPages As Collection
    m_CurrentType = "Text"
    m_LastStatus = 0
    Set txtPages = New Collection
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then
        m_LastStatus = 30
        Exit Function
    End If
    
    Dim content As String
    content = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    If Len(content) > 0 Then
        txtPages.Add content
    Else
        m_LastStatus = 31
    End If
    Set ExtractTextPages = txtPages
End Function

Private Function ExtractHtmlPages(ByVal filePath As String) As Collection
    Dim htmlPages As Collection
    m_CurrentType = "Html"
    m_LastStatus = 0
    Set htmlPages = New Collection
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then
        m_LastStatus = 40
        Exit Function
    End If
    
    Dim htmlContent As String
    htmlContent = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    Dim htmlDoc As Object
    On Error Resume Next
    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.body.innerHTML = htmlContent
    Dim text As String
    text = htmlDoc.body.innerText
    Set htmlDoc = Nothing
    On Error GoTo 0
    
    If Len(text) > 0 Then
        htmlPages.Add text
    Else
        m_LastStatus = 41
    End If
    Set ExtractHtmlPages = htmlPages
End Function

Private Function ExtractXmlPages(ByVal filePath As String) As Collection
    Dim xmlPages As Collection
    m_CurrentType = "Xml"
    m_LastStatus = 0
    Set xmlPages = New Collection
    Dim xmlDoc As Object
    
    On Error Resume Next
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePath
    On Error GoTo 0
    
    If xmlDoc Is Nothing Then
        m_LastStatus = 50
        Exit Function
    End If
    
    Dim text As String
    If Not xmlDoc.DocumentElement Is Nothing Then
        text = xmlDoc.DocumentElement.text
    End If
    Set xmlDoc = Nothing
    
    If Len(text) > 0 Then
        xmlPages.Add text
    Else
        m_LastStatus = 51
    End If
    Set ExtractXmlPages = xmlPages
End Function

Private Function ExtractImagePages(ByVal filePath As String) As Collection
    Dim imagePages As Collection
    m_CurrentType = "Image"
    m_LastStatus = 0
    Set imagePages = New Collection
    If m_WinOCR Is Nothing Then Set m_WinOCR = New WinOCR
    
    Dim vResult As Variant
    vResult = m_WinOCR.ImageToText(filePath, "", False)
    
    Dim text As String
    If IsArray(vResult) Then
        If UBound(vResult) >= 0 Then
            text = CStr(vResult(0))
        End If
    End If
    
    m_LastStatus = m_WinOCR.LastStatus
    
    If Len(text) > 0 Then
        imagePages.Add text
    ElseIf m_LastStatus = 0 Then
        m_LastStatus = 60
    End If
    Set ExtractImagePages = imagePages
End Function

Private Function ExtractMsgPages(ByVal filePath As String) As Collection
    Dim msgPages As Collection
    m_CurrentType = "Msg"
    m_LastStatus = 0
    Set msgPages = New Collection
    Dim oOutlook As Object
    Dim oItem As Object
    Dim bodyText As String
    Dim subject As String
    
    Set oOutlook = GetOutlookInstance()
    If oOutlook Is Nothing Then
        m_LastStatus = 70
        Exit Function
    End If
    
    On Error Resume Next
    Set oItem = oOutlook.session.OpenSharedItem(filePath)
    On Error GoTo 0
    
    If oItem Is Nothing Then
        m_LastStatus = 71
        Set oOutlook = Nothing
        Exit Function
    End If
    
    subject = oItem.subject
    If Len(oItem.body) > 0 Then
        bodyText = oItem.body
    ElseIf Len(oItem.HTMLBody) > 0 Then
        bodyText = StripHtmlTags(oItem.HTMLBody)
    End If
    
    If Len(subject) > 0 Then
        msgPages.Add subject
    End If
    
    msgPages.Add bodyText
    
    If m_ExtractAttachments And oItem.Attachments.count > 0 Then
        If m_CurrentDepth < MaxRecursionDepth Then
            m_CurrentDepth = m_CurrentDepth + 1
            Dim attTexts As Collection
            Set attTexts = ExtractMsgAttachmentsAsCollection(oItem)
            Dim attText As Variant
            For Each attText In attTexts
                msgPages.Add attText
            Next
            m_CurrentDepth = m_CurrentDepth - 1
        End If
    End If
    
    oItem.Close 0
    Set oItem = Nothing
    Set oOutlook = Nothing
    
    If msgPages.count <= 1 And Len(bodyText) = 0 Then m_LastStatus = 72
    Set ExtractMsgPages = msgPages
End Function

Private Function ExtractMsgAttachmentsAsCollection(ByVal oItem As Object) As Collection
    Dim attPages As Collection
    If m_CurrentDepth >= MaxRecursionDepth Then Exit Function
    m_CurrentDepth = m_CurrentDepth + 1
    Set attPages = New Collection
    Dim tempFolder As String
    tempFolder = GetTempFolder()
    
    Dim att As Object
    Dim tempPath As String
    Dim ext As String
    Dim attPages As Collection
    Dim page As Variant
    Dim prevType As String
    
    For Each att In oItem.Attachments
        ext = LCase(GetExtension(att.fileName))
        tempPath = tempFolder & "\" & NewGuid() & "." & ext
        
        On Error Resume Next
        att.SaveAsFile tempPath
        On Error GoTo 0
        
        If Len(Dir(tempPath)) > 0 Then
            prevType = m_CurrentType
            
            Set attPages = ExtractPages(tempPath)
            For Each page In attPages
                If Len(CStr(page)) > 0 Then
                    attPages.Add page
                End If
            Next
            
            m_CurrentType = prevType
            
            On Error Resume Next
            Kill tempPath
            On Error GoTo 0
        End If
    Next
    
    m_CurrentDepth = m_CurrentDepth - 1
    Set ExtractMsgAttachmentsAsCollection = attPages
End Function

Private Function ExtractEmlPages(ByVal filePath As String) As Collection
    Dim emlPages As Collection
    m_CurrentType = "Eml"
    m_LastStatus = 0
    Set emlPages = New Collection
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then
        m_LastStatus = 80
        Exit Function
    End If
    
    Dim content As String
    content = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    Dim subject As String
    Dim bodyText As String
    Dim contentType As String
    
    subject = ExtractEmlHeader(content, "Subject")
    bodyText = ExtractEmlBody(content, contentType)
    
    If Len(subject) > 0 Then
        emlPages.Add subject
    End If
    
    emlPages.Add bodyText
    
    If m_ExtractAttachments Then
        If m_CurrentDepth < MaxRecursionDepth Then
            m_CurrentDepth = m_CurrentDepth + 1
            Dim attTexts As Collection
            Set attTexts = ExtractEmlAttachmentsAsCollection(content)
            Dim attText As Variant
            For Each attText In attTexts
                emlPages.Add attText
            Next
            m_CurrentDepth = m_CurrentDepth - 1
        End If
    End If
    
    If emlPages.count <= 1 And Len(bodyText) = 0 Then m_LastStatus = 81
    Set ExtractEmlPages = emlPages
End Function

Private Function ExtractZipPages(ByVal filePath As String) As Collection
    Dim zipPages As Collection
    m_CurrentType = "Zip"
    m_LastStatus = 0
    Set zipPages = New Collection
    Dim sh As Object
    Dim zipFolder As Object
    Dim item As Object
    Dim tempFolder As String
    
    On Error Resume Next
    Set sh = CreateObject("Shell.Application")
    Set zipFolder = sh.Namespace(filePath)
    On Error GoTo 0
    
    If zipFolder Is Nothing Then
        m_LastStatus = 90
        Exit Function
    End If
    
    tempFolder = GetTempFolder()
    
    Dim retries As Long
    
    For Each item In zipFolder.Items
        Dim ext As String
        ext = LCase(GetExtension(item.Name))
        
        If ext = "zip" And m_CurrentDepth < MaxRecursionDepth Then
            Dim nestedZipPath As String
            nestedZipPath = tempFolder & "\" & item.Name
            
            Dim nestedExtracted As Boolean
            nestedExtracted = False
            
            For retries = 1 To 10
                On Error Resume Next
                zipFolder.CopyHere item.path, 16
                On Error GoTo 0
                
                On Error Resume Next
                Application.Wait Now + TimeValue("0:00:00.5")
                On Error GoTo 0
                
                If Len(Dir(nestedZipPath)) > 0 Then
                    nestedExtracted = True
                    Exit For
                End If
            Next
            
            If nestedExtracted Then
                m_CurrentDepth = m_CurrentDepth + 1
                Dim nestedPages As Collection
                Set nestedPages = ExtractZipPages(nestedZipPath)
                Dim nestedPage As Variant
                For Each nestedPage In nestedPages
                    If Len(CStr(nestedPage)) > 0 Then
                        zipPages.Add nestedPage
                    End If
                Next
                m_CurrentDepth = m_CurrentDepth - 1
                
                On Error Resume Next
                Kill nestedZipPath
                On Error GoTo 0
            End If
        Else
            Dim tempPath As String
            tempPath = tempFolder & "\" & item.Name
            
            Dim extracted As Boolean
            extracted = False
            
            For retries = 1 To 10
                On Error Resume Next
                zipFolder.CopyHere item.path, 16
                On Error GoTo 0
                
                On Error Resume Next
                Application.Wait Now + TimeValue("0:00:00.5")
                On Error GoTo 0
                
                If Len(Dir(tempPath)) > 0 Then
                    extracted = True
                    Exit For
                End If
            Next
            
            If extracted Then
                Dim prevType As String
                prevType = m_CurrentType
                
                Dim filePages As Collection
                Set filePages = ExtractPages(tempPath)
                Dim page As Variant
                For Each page In filePages
                    If Len(CStr(page)) > 0 Then
                        zipPages.Add page
                    End If
                Next
                
                m_CurrentType = prevType
                
                On Error Resume Next
                Kill tempPath
                On Error GoTo 0
            End If
        End If
    Next
    
    If zipPages.count = 0 Then m_LastStatus = 91
    Set ExtractZipPages = zipPages
End Function

Private Function ExtractRtfPages(ByVal filePath As String) As Collection
    Dim rtfPages As Collection
    m_CurrentType = "Rtf"
    m_LastStatus = 0
    Set rtfPages = New Collection
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then
        m_LastStatus = 100
        Exit Function
    End If
    
    Dim rtfContent As String
    On Error Resume Next
    rtfContent = file.ReadAll
    On Error GoTo 0
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    If Len(rtfContent) > 0 Then
        Dim htmlDoc As Object
        On Error Resume Next
        Set htmlDoc = CreateObject("htmlfile")
        htmlDoc.body.innerHTML = rtfContent
        Dim text As String
        text = htmlDoc.body.innerText
        Set htmlDoc = Nothing
        On Error GoTo 0
        
        If Len(text) > 0 Then
            rtfPages.Add text
        Else
            m_LastStatus = 101
        End If
    Else
        m_LastStatus = 101
    End If
    Set ExtractRtfPages = rtfPages
End Function

Private Function ExtractEmlHeader(ByVal content As String, ByVal headerName As String) As String
    ExtractEmlHeader = ""
    
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim i As Long
    Dim headerLine As String
    
    For i = LBound(lines) To UBound(lines)
        headerLine = LCase(lines(i))
        If Left(headerLine, Len(headerName) + 1) = LCase(headerName) & ":" Then
            ExtractEmlHeader = DecodeQuotedPrintable(Trim(Mid(lines(i), Len(headerName) + 2)))
            Exit Function
        End If
    Next i
End Function

Private Function ExtractEmlBody(ByVal content As String, ByRef contentTypeOut As String) As String
    ExtractEmlBody = ""
    contentTypeOut = ""
    
    Dim headerEnd As Long
    headerEnd = InStr(content, vbCrLf & vbCrLf)
    If headerEnd = 0 Then headerEnd = InStr(content, vbLf & vbLf)
    If headerEnd = 0 Then headerEnd = Len(content) + 1
    
    Dim headerPart As String
    Dim bodyPart As String
    headerPart = Left(content, headerEnd - 1)
    bodyPart = Mid(content, headerEnd + 2)
    
    Dim lines() As String
    lines = Split(headerPart, vbCrLf)
    
    Dim i As Long
    Dim line As String
    Dim contentType As String
    
    For i = LBound(lines) To UBound(lines)
        line = LCase(lines(i))
        If InStr(line, "content-type") > 0 Then
            contentType = lines(i)
            Exit For
        End If
    Next i
    
    contentTypeOut = contentType
    
    If InStr(contentType, "text/html") > 0 Then
        ExtractEmlBody = StripHtmlTags(DecodeQuotedPrintable(bodyPart))
    Else
        ExtractEmlBody = DecodeQuotedPrintable(bodyPart)
    End If
End Function

Private Function DecodeQuotedPrintable(ByVal text As String) As String
    DecodeQuotedPrintable = text
    
    If InStr(text, "=") = 0 Then Exit Function
    
    Dim result As String
    result = ""
    
    Dim i As Long
    Dim c As String
    Dim hexVal As String
    
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        
        If c = "=" Then
            If i + 2 <= Len(text) Then
                hexVal = Mid(text, i + 1, 2)
                If UCase(hexVal) = "0D0A" Or hexVal = "0D" Or hexVal = "0A" Then
                    i = i + 2
                    If UCase(hexVal) = "0D0A" Then i = i + 1
                    result = result & vbCrLf
                ElseIf IsNumeric("&H" & hexVal) Then
                    result = result & Chr(CLng("&H" & hexVal))
                    i = i + 2
                Else
                    result = result & c
                End If
            Else
                result = result & c
            End If
        Else
            result = result & c
        End If
    Next i
    
    DecodeQuotedPrintable = result
End Function

Private Function ExtractEmlAttachmentsAsCollection(ByVal emlContent As String) As Collection
    Dim attPages As Collection
    If m_CurrentDepth >= MaxRecursionDepth Then Exit Function
    m_CurrentDepth = m_CurrentDepth + 1
    Set attPages = New Collection
    Dim boundary As String
    boundary = ExtractBoundary(emlContent)
    
    If Len(boundary) = 0 Then
        m_CurrentDepth = m_CurrentDepth - 1
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(emlContent, "--" & boundary)
    
    Dim i As Long
    Dim part As String
    Dim fileName As String
    Dim decodedBytes() As Byte
    Dim tempPath As String
    Dim attPages As Collection
    Dim page As Variant
    Dim prevType As String
    Dim ext As String
    Dim fNum As Integer
    
    For i = LBound(parts) To UBound(parts)
        part = parts(i)
        
        fileName = ExtractFilename(part)
        If Len(fileName) > 0 Then
            decodedBytes = DecodeBase64Attachment(part)
            
            If UBound(decodedBytes) >= 0 Then
                ext = LCase(GetExtension(fileName))
                If Len(ext) = 0 Then ext = "bin"
                tempPath = GetTempFolder() & "\" & NewGuid() & "." & ext
                
                On Error Resume Next
                fNum = FreeFile
                Open tempPath For Binary As #fNum
                Put #fNum, , decodedBytes
                Close #fNum
                On Error GoTo 0
                
                If Len(Dir(tempPath)) > 0 Then
                    prevType = m_CurrentType
                    
                    Set attPages = ExtractPages(tempPath)
                    For Each page In attPages
                        If Len(CStr(page)) > 0 Then
                            attPages.Add page
                        End If
                    Next
                    
                    m_CurrentType = prevType
                    
                    On Error Resume Next
                    Kill tempPath
                    On Error GoTo 0
                End If
            End If
        End If
    Next i
    
    m_CurrentDepth = m_CurrentDepth - 1
    Set ExtractEmlAttachmentsAsCollection = attPages
End Function

Private Function ExtractBoundary(ByVal content As String) As String
    ExtractBoundary = ""
    
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim i As Long
    Dim line As String
    
    For i = LBound(lines) To UBound(lines)
        line = LCase(lines(i))
        If InStr(line, "boundary") > 0 Then
            Dim parts2() As String
            parts2 = Split(line, "=")
            If UBound(parts2) >= 1 Then
                ExtractBoundary = Trim(Replace(parts2(1), """", ""))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ExtractFilename(ByVal part As String) As String
    ExtractFilename = ""
    
    Dim lines() As String
    lines = Split(part, vbCrLf)
    
    Dim i As Long
    Dim line As String
    
    For i = LBound(lines) To UBound(lines)
        line = LCase(lines(i))
        If InStr(line, "filename") > 0 Then
            Dim parts2() As String
            parts2 = Split(line, "=")
            If UBound(parts2) >= 1 Then
                ExtractFilename = Trim(Replace(parts2(1), """", ""))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function DecodeBase64Attachment(ByVal part As String) As Byte()
    On Error Resume Next
    
    Dim base64Start As Long
    base64Start = InStr(part, vbCrLf & vbCrLf)
    If base64Start = 0 Then base64Start = InStr(part, vbLf & vbLf)
    If base64Start = 0 Then Exit Function
    
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    
    Dim node As Object
    Set node = xmlDoc.createElement("b64")
    node.DataType = "bin.base64"
    node.text = Mid(part, base64Start + 4)
    
    DecodeBase64Attachment = node.NodeTypedValue
    
    Set node = Nothing
    Set xmlDoc = Nothing
    
    On Error GoTo 0
End Function

Private Function StripHtmlTags(ByVal html As String) As String
    StripHtmlTags = html
    
    Dim htmlDoc As Object
    On Error Resume Next
    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.body.innerHTML = html
    StripHtmlTags = htmlDoc.body.innerText
    Set htmlDoc = Nothing
    On Error GoTo 0
End Function

