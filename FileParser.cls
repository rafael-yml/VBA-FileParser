VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MaxRecursionDepth As Long = 5

Private m_PdfParser As PdfParser
Private m_CurrentType As String
Private m_CurrentDepth As Long
Private m_ExtractAttachments As Boolean

Public Property Get Parser() As PdfParser
    If m_PdfParser Is Nothing Then
        Set m_PdfParser = New PdfParser
    End If
    Set Parser = m_PdfParser
End Property

Public Property Get LastType() As String
    LastType = m_CurrentType
End Property

Public Function ExtractText(ByVal filePath As String, Optional ByVal extractAttachments As Boolean = True) As String
    m_CurrentDepth = 0
    m_ExtractAttachments = extractAttachments
    m_CurrentType = ""
    
    Dim ext As String
    ext = LCase(GetExtension(filePath))
    
    Select Case ext
        Case "pdf"
            ExtractText = ExtractPdfText(filePath)
        Case "xls", "xlsx", "xlsb", "xlsm"
            ExtractText = ExtractExcelText(filePath)
        Case "doc", "docx", "docm"
            ExtractText = ExtractWordText(filePath)
        Case "txt", "csv"
            ExtractText = ExtractTextFile(filePath)
        Case "html", "htm"
            ExtractText = ExtractHtmlText(filePath)
        Case "xml"
            ExtractText = ExtractXmlText(filePath)
        Case "png", "jpg", "jpeg", "gif", "bmp", "webp", "tiff", "tif"
            ExtractText = ExtractImageText(filePath)
        Case "msg"
            ExtractText = ExtractMsgText(filePath)
        Case "eml"
            ExtractText = ExtractEmlText(filePath)
        Case Else
            m_CurrentType = "Unknown"
            ExtractText = ""
    End Select
End Function

Public Function ExtractPages(ByVal filePath As String) As Collection
    Set ExtractPages = New Collection
    m_CurrentType = ""
    
    Dim ext As String
    ext = LCase(GetExtension(filePath))
    
    Select Case ext
        Case "pdf"
            ExtractPages = ExtractPdfPages(filePath)
        Case "xls", "xlsx", "xlsb", "xlsm"
            ExtractPages = ExtractExcelPages(filePath)
        Case "doc", "docx", "docm"
            ExtractPages = ExtractWordPages(filePath)
        Case "txt", "csv"
            ExtractPages = ExtractTextPages(filePath)
        Case "html", "htm"
            ExtractPages = ExtractHtmlPages(filePath)
        Case "xml"
            ExtractPages = ExtractXmlPages(filePath)
        Case "png", "jpg", "jpeg", "gif", "bmp", "webp", "tiff", "tif"
            ExtractPages = ExtractImagePages(filePath)
        Case "msg"
            ExtractPages = ExtractMsgPages(filePath)
        Case "eml"
            ExtractPages = ExtractEmlPages(filePath)
    End Select
End Function

Private Function GetExtension(ByVal filePath As String) As String
    Dim parts() As String
    parts = Split(filePath, ".")
    If UBound(parts) >= 0 Then
        GetExtension = parts(UBound(parts))
    Else
        GetExtension = ""
    End If
End Function

Private Function GetTempFolder() As String
    GetTempFolder = Environ("TEMP") & "\FileParser_" & Format(Now, "YYYYMMDDHHMMSS")
End Function

Private Sub EnsureTempFolder(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function ExtractPdfText(ByVal filePath As String) As String
    m_CurrentType = "Pdf"
    ExtractPdfText = Parser.ExtractText(filePath)
End Function

Private Function ExtractPdfPages(ByVal filePath As String) As Collection
    m_CurrentType = "Pdf"
    Set ExtractPdfPages = Parser.ExtractPages(filePath)
End Function

Private Function ExtractExcelText(ByVal filePath As String) As String
    m_CurrentType = "Excel"
    ExtractExcelText = ExtractExcelTextCore(filePath)
    
    If Len(ExtractExcelText) = 0 And m_CurrentDepth < MaxRecursionDepth Then
        ExtractExcelText = FallbackWordPdfToOCR(filePath)
    End If
End Function

Private Function ExtractExcelTextCore(ByVal filePath As String) As String
    ExtractExcelTextCore = ""
    
    Dim wb As Object
    Dim ws As Object
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim cellValue As Variant
    
    On Error Resume Next
    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    On Error GoTo 0
    
    If wb Is Nothing Then Exit Function
    
    On Error Resume Next
    Application.ScreenUpdating = False
    On Error GoTo 0
    
    Set ws = wb.ActiveSheet
    On Error Resume Next
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    lastCol = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
    On Error GoTo 0
    
    If lastRow = 0 Then lastRow = 1
    If lastCol = 0 Then lastCol = 1
    
    For r = 1 To lastRow
        For c = 1 To lastCol
            On Error Resume Next
            cellValue = ws.Cells(r, c).Value
            On Error GoTo 0
            If Not IsEmpty(cellValue) Then
                ExtractExcelTextCore = ExtractExcelTextCore & CStr(cellValue) & " "
            End If
        Next c
        If lastCol > 0 Then ExtractExcelTextCore = ExtractExcelTextCore & vbCrLf
    Next r
    
    Application.ScreenUpdating = True
    wb.Close SaveChanges:=False
    Set wb = Nothing
End Function

Private Function ExtractExcelPages(ByVal filePath As String) As Collection
    Set ExtractExcelPages = New Collection
    
    Dim wb As Object
    Dim ws As Object
    Dim sheetIndex As Long
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim cellValue As Variant
    Dim sheetText As String
    
    On Error Resume Next
    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    On Error GoTo 0
    
    If wb Is Nothing Then Exit Function
    
    On Error Resume Next
    Application.ScreenUpdating = False
    On Error GoTo 0
    
    For sheetIndex = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(sheetIndex)
        sheetText = ""
        
        On Error Resume Next
        lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
        lastCol = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
        On Error GoTo 0
        
        If lastRow = 0 Then lastRow = 1
        If lastCol = 0 Then lastCol = 1
        
        For r = 1 To lastRow
            For c = 1 To lastCol
                On Error Resume Next
                cellValue = ws.Cells(r, c).Value
                On Error GoTo 0
                If Not IsEmpty(cellValue) Then
                    sheetText = sheetText & CStr(cellValue) & " "
                End If
            Next c
            If lastCol > 0 Then sheetText = sheetText & vbCrLf
        Next r
        
        If Len(sheetText) > 0 Then
            ExtractExcelPages.Add sheetText
        End If
    Next sheetIndex
    
    Application.ScreenUpdating = True
    wb.Close SaveChanges:=False
    Set wb = Nothing
    
    If ExtractExcelPages.Count = 0 Then m_CurrentType = ""
End Function

Private Function ExtractWordText(ByVal filePath As String) As String
    m_CurrentType = "Word"
    ExtractWordText = ExtractWordTextCore(filePath)
    
    If Len(ExtractWordText) = 0 And m_CurrentDepth < MaxRecursionDepth Then
        ExtractWordText = FallbackWordPdfToOCR(filePath)
    End If
End Function

Private Function ExtractWordTextCore(ByVal filePath As String) As String
    ExtractWordTextCore = ""
    
    Dim oWord As Object
    Dim oDoc As Object
    
    On Error Resume Next
    Set oWord = CreateObject("Word.Application")
    On Error GoTo 0
    
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
    
    ExtractWordTextCore = oDoc.Content.Text
    
    oDoc.Close SaveChanges:=False
    oWord.Quit
    Set oDoc = Nothing
    Set oWord = Nothing
End Function

Private Function ExtractWordPages(ByVal filePath As String) As Collection
    Set ExtractWordPages = New Collection
    m_CurrentType = "Word"
    
    Dim oWord As Object
    Dim oDoc As Object
    Dim i As Long
    Dim pageText As String
    
    On Error Resume Next
    Set oWord = CreateObject("Word.Application")
    On Error GoTo 0
    
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
    
    For i = 1 To oDoc.ComputeStatistics(2)
        On Error Resume Next
        oDoc.Bookmarks("\Page").Range.Select
        pageText = oWord.Selection.Text
        On Error GoTo 0
        If Len(pageText) > 0 Then
            ExtractWordPages.Add pageText
        End If
    Next i
    
    If ExtractWordPages.Count = 0 Then
        ExtractWordPages.Add oDoc.Content.Text
    End If
    
    oDoc.Close SaveChanges:=False
    oWord.Quit
    Set oDoc = Nothing
    Set oWord = Nothing
End Function

Private Function FallbackWordPdfToOCR(ByVal filePath As String) As String
    FallbackWordPdfToOCR = ""
    
    Dim tempPdf As String
    tempPdf = GetTempFolder() & "_fallback.pdf"
    EnsureTempFolder GetTempFolder()
    
    Dim oWord As Object
    Dim oDoc As Object
    
    On Error Resume Next
    Set oWord = CreateObject("Word.Application")
    On Error GoTo 0
    
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
        FallbackWordPdfToOCR = Parser.ExtractText(tempPdf)
        On Error Resume Next
        Kill tempPdf
        On Error GoTo 0
    End If
End Function

Private Function ExtractTextFile(ByVal filePath As String) As String
    m_CurrentType = "Text"
    
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then Exit Function
    
    ExtractTextFile = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
End Function

Private Function ExtractTextPages(ByVal filePath As String) As Collection
    Set ExtractTextPages = New Collection
    m_CurrentType = "Text"
    
    Dim content As String
    content = ExtractTextFile(filePath)
    
    If Len(content) > 0 Then
        ExtractTextPages.Add content
    End If
End Function

Private Function ExtractHtmlText(ByVal filePath As String) As String
    m_CurrentType = "Html"
    
    Dim fso As Object
    Dim file As Object
    Dim htmlDoc As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then Exit Function
    
    Dim htmlContent As String
    htmlContent = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    On Error Resume Next
    Set htmlDoc = CreateObject("MSHTML.HTMLFile")
    htmlDoc.Body.innerHTML = htmlContent
    ExtractHtmlText = htmlDoc.Body.innerText
    Set htmlDoc = Nothing
    On Error GoTo 0
End Function

Private Function ExtractHtmlPages(ByVal filePath As String) As Collection
    Set ExtractHtmlPages = New Collection
    m_CurrentType = "Html"
    
    Dim content As String
    content = ExtractHtmlText(filePath)
    
    If Len(content) > 0 Then
        ExtractHtmlPages.Add content
    End If
End Function

Private Function ExtractXmlText(ByVal filePath As String) As String
    m_CurrentType = "Xml"
    
    Dim xmlDoc As Object
    
    On Error Resume Next
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePath
    On Error GoTo 0
    
    If Not xmlDoc Is Nothing Then
        If Not xmlDoc.DocumentElement Is Nothing Then
            ExtractXmlText = xmlDoc.DocumentElement.Text
        End If
        Set xmlDoc = Nothing
    End If
End Function

Private Function ExtractXmlPages(ByVal filePath As String) As Collection
    Set ExtractXmlPages = New Collection
    m_CurrentType = "Xml"
    
    Dim content As String
    content = ExtractXmlText(filePath)
    
    If Len(content) > 0 Then
        ExtractXmlPages.Add content
    End If
End Function

Private Function ExtractImageText(ByVal filePath As String) As String
    m_CurrentType = "Image"
    
    Dim ocr As Object
    On Error Resume Next
    Set ocr = New WinOCR
    On Error GoTo 0
    
    If ocr Is Nothing Then Exit Function
    
    Dim aPage() As Byte
    Dim ext As String
    ext = LCase(GetExtension(filePath))
    
    On Error Resume Next
    Open filePath For Binary As #1
        ReDim aPage(0 To LOF(1) - 1)
        Get #1, , aPage
    Close #1
    On Error GoTo 0
    
    ExtractImageText = ocr.BytesToText(aPage)
    Set ocr = Nothing
End Function

Private Function ExtractImagePages(ByVal filePath As String) As Collection
    Set ExtractImagePages = New Collection
    m_CurrentType = "Image"
    
    Dim content As String
    content = ExtractImageText(filePath)
    
    If Len(content) > 0 Then
        ExtractImagePages.Add content
    End If
End Function

Private Function ExtractMsgText(ByVal filePath As String) As String
    m_CurrentType = "Msg"
    
    Dim oOutlook As Object
    Dim oItem As Object
    Dim bodyText As String
    
    On Error Resume Next
    Set oOutlook = CreateObject("Outlook.Application")
    Set oItem = oOutlook.Session.OpenSharedItem(filePath)
    On Error GoTo 0
    
    If oItem Is Nothing Then Exit Function
    
    If Len(oItem.Body) > 0 Then
        bodyText = oItem.Body
    ElseIf Len(oItem.HTMLBody) > 0 Then
        bodyText = StripHtmlTags(oItem.HTMLBody)
    End If
    
    ExtractMsgText = bodyText
    
    If m_ExtractAttachments And oItem.Attachments.Count > 0 Then
        ExtractMsgText = ExtractMsgText & vbCrLf & ExtractMsgAttachments(oItem)
    End If
    
    oItem.Close 0
    Set oItem = Nothing
    Set oOutlook = Nothing
End Function

Private Function ExtractMsgPages(ByVal filePath As String) As Collection
    Set ExtractMsgPages = New Collection
    m_CurrentType = "Msg"
    
    Dim oOutlook As Object
    Dim oItem As Object
    Dim bodyText As String
    
    On Error Resume Next
    Set oOutlook = CreateObject("Outlook.Application")
    Set oItem = oOutlook.Session.OpenSharedItem(filePath)
    On Error GoTo 0
    
    If oItem Is Nothing Then Exit Function
    
    If Len(oItem.Body) > 0 Then
        bodyText = oItem.Body
    ElseIf Len(oItem.HTMLBody) > 0 Then
        bodyText = StripHtmlTags(oItem.HTMLBody)
    End If
    
    ExtractMsgPages.Add bodyText
    
    If m_ExtractAttachments And oItem.Attachments.Count > 0 Then
        Dim attTexts As Collection
        Set attTexts = ExtractMsgAttachmentsAsCollection(oItem)
        Dim attText As Variant
        For Each attText In attTexts
            ExtractMsgPages.Add attText
        Next
    End If
    
    oItem.Close 0
    Set oItem = Nothing
    Set oOutlook = Nothing
End Function

Private Function ExtractMsgAttachments(ByVal oItem As Object) As String
    ExtractMsgAttachments = ""
    
    Dim tempFolder As String
    tempFolder = GetTempFolder()
    EnsureTempFolder tempFolder
    
    m_CurrentDepth = m_CurrentDepth + 1
    
    If m_CurrentDepth > MaxRecursionDepth Then
        m_CurrentDepth = m_CurrentDepth - 1
        Exit Function
    End If
    
    Dim att As Object
    Dim tempPath As String
    Dim attText As String
    Dim prevType As String
    
    For Each att In oItem.Attachments
        tempPath = tempFolder & "\" & att.FileName
        On Error Resume Next
        att.SaveAsFile tempPath
        On Error GoTo 0
        
        If Len(Dir(tempPath)) > 0 Then
            prevType = m_CurrentType
            attText = ExtractText(tempPath, m_ExtractAttachments)
            If Len(attText) > 0 Then
                ExtractMsgAttachments = ExtractMsgAttachments & vbCrLf & attText
            End If
            m_CurrentType = prevType
            
            On Error Resume Next
            Kill tempPath
            On Error GoTo 0
        End If
    Next
    
    On Error Resume Next
    RmDir tempFolder
    On Error GoTo 0
    
    m_CurrentDepth = m_CurrentDepth - 1
End Function

Private Function ExtractMsgAttachmentsAsCollection(ByVal oItem As Object) As Collection
    Set ExtractMsgAttachmentsAsCollection = New Collection
    
    Dim tempFolder As String
    tempFolder = GetTempFolder()
    EnsureTempFolder tempFolder
    
    m_CurrentDepth = m_CurrentDepth + 1
    
    If m_CurrentDepth > MaxRecursionDepth Then
        m_CurrentDepth = m_CurrentDepth - 1
        Exit Function
    End If
    
    Dim att As Object
    Dim tempPath As String
    Dim attText As String
    Dim prevType As String
    Dim attPages As Collection
    Dim page As Variant
    
    For Each att In oItem.Attachments
        tempPath = tempFolder & "\" & att.FileName
        On Error Resume Next
        att.SaveAsFile tempPath
        On Error GoTo 0
        
        If Len(Dir(tempPath)) > 0 Then
            prevType = m_CurrentType
            
            Set attPages = ExtractPages(tempPath)
            For Each page In attPages
                If Len(CStr(page)) > 0 Then
                    ExtractMsgAttachmentsAsCollection.Add page
                End If
            Next
            
            m_CurrentType = prevType
            
            On Error Resume Next
            Kill tempPath
            On Error GoTo 0
        End If
    Next
    
    On Error Resume Next
    RmDir tempFolder
    On Error GoTo 0
    
    m_CurrentDepth = m_CurrentDepth - 1
End Function

Private Function ExtractEmlText(ByVal filePath As String) As String
    m_CurrentType = "Eml"
    
    Dim fso As Object
    Dim file As Object
    Dim lines() As String
    Dim i As Long
    Dim inBody As Boolean
    Dim bodyStart As Long
    Dim headerEnd As Long
    Dim contentType As String
    Dim charset As String
    Dim htmlBody As String
    Dim plainBody As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then Exit Function
    
    Dim content As String
    content = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    lines = Split(content, vbCrLf)
    
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) = 0 Then
            headerEnd = i
            Exit For
        End If
        If LCase(Left(lines(i), 13)) = "content-type" Then
            contentType = lines(i)
        End If
    Next i
    
    If headerEnd = 0 Then headerEnd = UBound(lines)
    
    For i = headerEnd + 1 To UBound(lines)
        plainBody = plainBody & lines(i) & vbCrLf
    Next i
    
    If InStr(contentType, "text/html") > 0 Then
        htmlBody = Mid(content, InStr(content, vbCrLf & vbCrLf) + 4)
        ExtractEmlText = StripHtmlTags(htmlBody)
    Else
        ExtractEmlText = plainBody
    End If
    
    If m_ExtractAttachments Then
        ExtractEmlText = ExtractEmlText & vbCrLf & ExtractEmlAttachments(content)
    End If
End Function

Private Function ExtractEmlPages(ByVal filePath As String) As Collection
    Set ExtractEmlPages = New Collection
    m_CurrentType = "Eml"
    
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set file = fso.OpenTextFile(filePath, 1)
    On Error GoTo 0
    
    If file Is Nothing Then Exit Function
    
    Dim content As String
    content = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    Dim bodyText As String
    bodyText = ExtractEmlText(filePath)
    ExtractEmlPages.Add bodyText
    
    If m_ExtractAttachments Then
        Dim attTexts As Collection
        Set attTexts = ExtractEmlAttachmentsAsCollection(content)
        Dim attText As Variant
        For Each attText In attTexts
            ExtractEmlPages.Add attText
        Next
    End If
End Function

Private Function ExtractEmlAttachments(ByVal emlContent As String) As String
    ExtractEmlAttachments = ""
    
    Dim tempFolder As String
    tempFolder = GetTempFolder()
    EnsureTempFolder tempFolder
    
    m_CurrentDepth = m_CurrentDepth + 1
    
    If m_CurrentDepth > MaxRecursionDepth Then
        m_CurrentDepth = m_CurrentDepth - 1
        Exit Function
    End If
    
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
    Dim filename As String
    Dim decodedBytes() As Byte
    Dim tempPath As String
    Dim attText As String
    Dim prevType As String
    Dim fileNum As Long
    
    For i = LBound(parts) To UBound(parts)
        part = parts(i)
        
        filename = ExtractFilename(part)
        If Len(filename) > 0 Then
            decodedBytes = DecodeBase64Attachment(part)
            
            If UBound(decodedBytes) >= 0 Then
                tempPath = tempFolder & "\" & filename
                
                On Error Resume Next
                Open tempPath For Binary As #1
                    Put #1, , decodedBytes
                Close #1
                On Error GoTo 0
                
                If Len(Dir(tempPath)) > 0 Then
                    prevType = m_CurrentType
                    attText = ExtractText(tempPath, m_ExtractAttachments)
                    If Len(attText) > 0 Then
                        ExtractEmlAttachments = ExtractEmlAttachments & vbCrLf & attText
                    End If
                    m_CurrentType = prevType
                    
                    On Error Resume Next
                    Kill tempPath
                    On Error GoTo 0
                End If
            End If
        End If
    Next i
    
    On Error Resume Next
    RmDir tempFolder
    On Error GoTo 0
    
    m_CurrentDepth = m_CurrentDepth - 1
End Function

Private Function ExtractEmlAttachmentsAsCollection(ByVal emlContent As String) As Collection
    Set ExtractEmlAttachmentsAsCollection = New Collection
    
    Dim tempFolder As String
    tempFolder = GetTempFolder()
    EnsureTempFolder tempFolder
    
    m_CurrentDepth = m_CurrentDepth + 1
    
    If m_CurrentDepth > MaxRecursionDepth Then
        m_CurrentDepth = m_CurrentDepth - 1
        Exit Function
    End If
    
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
    Dim filename As String
    Dim decodedBytes() As Byte
    Dim tempPath As String
    Dim attPages As Collection
    Dim page As Variant
    Dim prevType As String
    
    For i = LBound(parts) To UBound(parts)
        part = parts(i)
        
        filename = ExtractFilename(part)
        If Len(filename) > 0 Then
            decodedBytes = DecodeBase64Attachment(part)
            
            If UBound(decodedBytes) >= 0 Then
                tempPath = tempFolder & "\" & filename
                
                On Error Resume Next
                Open tempPath For Binary As #1
                    Put #1, , decodedBytes
                Close #1
                On Error GoTo 0
                
                If Len(Dir(tempPath)) > 0 Then
                    prevType = m_CurrentType
                    Set attPages = ExtractPages(tempPath)
                    For Each page In attPages
                        If Len(CStr(page)) > 0 Then
                            ExtractEmlAttachmentsAsCollection.Add page
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
    
    On Error Resume Next
    RmDir tempFolder
    On Error GoTo 0
    
    m_CurrentDepth = m_CurrentDepth - 1
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
    
    If base64Start = 0 Then
        base64Start = InStr(part, vbLf & vbLf)
    End If
    
    If base64Start = 0 Then Exit Function
    
    Dim encodedPart As String
    encodedPart = Trim(Mid(part, base64Start + 4))
    
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    
    Dim node As Object
    Set node = xmlDoc.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = encodedPart
    
    DecodeBase64Attachment = node.NodeTypedValue
    
    Set node = Nothing
    Set xmlDoc = Nothing
    
    On Error GoTo 0
End Function

Private Function StripHtmlTags(ByVal html As String) As String
    StripHtmlTags = html
    
    Dim htmlDoc As Object
    On Error Resume Next
    Set htmlDoc = CreateObject("MSHTML.HTMLFile")
    htmlDoc.Body.innerHTML = html
    StripHtmlTags = htmlDoc.Body.innerText
    Set htmlDoc = Nothing
    On Error GoTo 0
End Function

Private Sub Class_Initialize()
    m_CurrentType = ""
    m_CurrentDepth = 0
    m_ExtractAttachments = True
End Sub