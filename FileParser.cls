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
    
    If pages.Count > 0 Then
        Dim i As Long
        For i = 1 To pages.Count
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
    Set ExtractPages = New Collection
    m_LastStatus = 0
    m_CurrentDepth = 0
    m_ExtractAttachments = extractAttachments
    m_CurrentType = ""
    m_TempFolder = ""
    
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
            Set ExtractPages = ExtractPdfPages(filePath)
        Case "xls", "xlsx", "xlsb", "xlsm"
            Set ExtractPages = ExtractExcelPages(filePath)
        Case "doc", "docx", "docm"
            Set ExtractPages = ExtractWordPages(filePath)
        Case "txt", "csv"
            Set ExtractPages = ExtractTextPages(filePath)
        Case "html", "htm"
            Set ExtractPages = ExtractHtmlPages(filePath)
        Case "xml"
            Set ExtractPages = ExtractXmlPages(filePath)
        Case "png", "jpg", "jpeg", "gif", "bmp", "webp", "tiff", "tif"
            Set ExtractPages = ExtractImagePages(filePath)
        Case "msg"
            Set ExtractPages = ExtractMsgPages(filePath)
        Case "eml"
            Set ExtractPages = ExtractEmlPages(filePath)
    End Select
    
    If ExtractPages.Count = 0 And m_LastStatus = 0 Then
        m_LastStatus = 99
    End If
        
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
    
    If Not m_Outlook Is Nothing Then
        On Error Resume Next
        m_Outlook.Quit
        Set m_Outlook = Nothing
        On Error GoTo 0
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
    Set ExtractPdfPages = New Collection
    m_CurrentType = "Pdf"
    Set ExtractPdfPages = Parser.ExtractPages(filePath)
    m_LastStatus = Parser.LastStatus
End Function

Private Function ExtractExcelPages(ByVal filePath As String) As Collection
    Set ExtractExcelPages = New Collection
    m_CurrentType = "Excel"
    m_LastStatus = 0
    
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
    
    If wb Is Nothing Then
        m_LastStatus = 10
        Set ExtractExcelPages = FallbackToPdf(filePath)
        Exit Function
    End If
    
    For sheetIndex = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(sheetIndex)
        sheetText = ""
        
        On Error Resume Next
        lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).row
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
    
    wb.Close SaveChanges:=False
    Set wb = Nothing
    
    If ExtractExcelPages.Count = 0 And m_CurrentDepth < MaxRecursionDepth Then
        Set ExtractExcelPages = FallbackToPdf(filePath)
    End If
    
    If ExtractExcelPages.Count = 0 Then m_LastStatus = 11
End Function

Private Function ExtractWordPages(ByVal filePath As String) As Collection
    Set ExtractWordPages = New Collection
    m_CurrentType = "Word"
    m_LastStatus = 0
    
    If m_WdCOM Is Nothing Then Set m_WdCOM = New WdCOM
    
    Dim text As String
    text = m_WdCOM.ExtractText(filePath)
    m_LastStatus = m_WdCOM.LastStatus
    
    If Len(text) > 0 Then
        ExtractWordPages.Add text
    End If
    
    If ExtractWordPages.Count = 0 And m_CurrentDepth < MaxRecursionDepth Then
        Set ExtractWordPages = FallbackToPdf(filePath)
    End If
    
    If ExtractWordPages.Count = 0 Then m_LastStatus = 21
End Function

Private Function FallbackToPdf(ByVal filePath As String) As Collection
    Set FallbackToPdf = New Collection
    
    m_CurrentDepth = m_CurrentDepth + 1
    If m_CurrentDepth > MaxRecursionDepth Then
        m_CurrentDepth = m_CurrentDepth - 1
        Exit Function
    End If
    
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
        Set FallbackToPdf = Parser.ExtractPages(tempPdf)
        m_CurrentType = "Pdf"
        On Error Resume Next
        Kill tempPdf
        On Error GoTo 0
    End If
    
    m_CurrentDepth = m_CurrentDepth - 1
End Function

Private Function ExtractTextPages(ByVal filePath As String) As Collection
    Set ExtractTextPages = New Collection
    m_CurrentType = "Text"
    m_LastStatus = 0
    
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
        ExtractTextPages.Add content
    Else
        m_LastStatus = 31
    End If
End Function

Private Function ExtractHtmlPages(ByVal filePath As String) As Collection
    Set ExtractHtmlPages = New Collection
    m_CurrentType = "Html"
    m_LastStatus = 0
    
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
    Set htmlDoc = CreateObject("MSHTML.HTMLFile")
    htmlDoc.Body.innerHTML = htmlContent
    Dim text As String
    text = htmlDoc.Body.innerText
    Set htmlDoc = Nothing
    On Error GoTo 0
    
    If Len(text) > 0 Then
        ExtractHtmlPages.Add text
    Else
        m_LastStatus = 41
    End If
End Function

Private Function ExtractXmlPages(ByVal filePath As String) As Collection
    Set ExtractXmlPages = New Collection
    m_CurrentType = "Xml"
    m_LastStatus = 0
    
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
        ExtractXmlPages.Add text
    Else
        m_LastStatus = 51
    End If
End Function

Private Function ExtractImagePages(ByVal filePath As String) As Collection
    Set ExtractImagePages = New Collection
    m_CurrentType = "Image"
    m_LastStatus = 0
    
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
        ExtractImagePages.Add text
    ElseIf m_LastStatus = 0 Then
        m_LastStatus = 60
    End If
End Function

Private Function ExtractMsgPages(ByVal filePath As String) As Collection
    Set ExtractMsgPages = New Collection
    m_CurrentType = "Msg"
    m_LastStatus = 0
    
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
    Set oItem = oOutlook.Session.OpenSharedItem(filePath)
    On Error GoTo 0
    
    If oItem Is Nothing Then
        m_LastStatus = 71
        Set oOutlook = Nothing
        Exit Function
    End If
    
    subject = oItem.subject
    If Len(oItem.Body) > 0 Then
        bodyText = oItem.Body
    ElseIf Len(oItem.HTMLBody) > 0 Then
        bodyText = StripHtmlTags(oItem.HTMLBody)
    End If
    
    If Len(subject) > 0 Then
        ExtractMsgPages.Add subject
    End If
    
    ExtractMsgPages.Add bodyText
    
    If m_ExtractAttachments And oItem.Attachments.Count > 0 Then
        If m_CurrentDepth < MaxRecursionDepth Then
            m_CurrentDepth = m_CurrentDepth + 1
            Dim attTexts As Collection
            Set attTexts = ExtractMsgAttachmentsAsCollection(oItem)
            Dim attText As Variant
            For Each attText In attTexts
                ExtractMsgPages.Add attText
            Next
            m_CurrentDepth = m_CurrentDepth - 1
        End If
    End If
    
    oItem.Close 0
    Set oItem = Nothing
    Set oOutlook = Nothing
    
    If ExtractMsgPages.Count <= 1 And Len(bodyText) = 0 Then m_LastStatus = 72
End Function

Private Function ExtractMsgAttachmentsAsCollection(ByVal oItem As Object) As Collection
    Set ExtractMsgAttachmentsAsCollection = New Collection
    
    If m_CurrentDepth >= MaxRecursionDepth Then Exit Function
    
    m_CurrentDepth = m_CurrentDepth + 1
    
    Dim tempFolder As String
    tempFolder = GetTempFolder()
    
    Dim att As Object
    Dim tempPath As String
    Dim ext As String
    Dim attPages As Collection
    Dim page As Variant
    Dim prevType As String
    
    For Each att In oItem.Attachments
        ext = LCase(GetExtension(att.filename))
        tempPath = tempFolder & "\" & NewGuid() & "." & ext
        
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
    
    m_CurrentDepth = m_CurrentDepth - 1
End Function

Private Function ExtractEmlPages(ByVal filePath As String) As Collection
    Set ExtractEmlPages = New Collection
    m_CurrentType = "Eml"
    m_LastStatus = 0
    
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
        ExtractEmlPages.Add subject
    End If
    
    ExtractEmlPages.Add bodyText
    
    If m_ExtractAttachments Then
        If m_CurrentDepth < MaxRecursionDepth Then
            m_CurrentDepth = m_CurrentDepth + 1
            Dim attTexts As Collection
            Set attTexts = ExtractEmlAttachmentsAsCollection(content)
            Dim attText As Variant
            For Each attText In attTexts
                ExtractEmlPages.Add attText
            Next
            m_CurrentDepth = m_CurrentDepth - 1
        End If
    End If
    
    If ExtractEmlPages.Count <= 1 And Len(bodyText) = 0 Then m_LastStatus = 81
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
    Set ExtractEmlAttachmentsAsCollection = New Collection
    
    If m_CurrentDepth >= MaxRecursionDepth Then Exit Function
    
    m_CurrentDepth = m_CurrentDepth + 1
    
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
    Dim ext As String
    Dim fNum As Integer
    
    For i = LBound(parts) To UBound(parts)
        part = parts(i)
        
        filename = ExtractFilename(part)
        If Len(filename) > 0 Then
            decodedBytes = DecodeBase64Attachment(part)
            
            If UBound(decodedBytes) >= 0 Then
                ext = LCase(GetExtension(filename))
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
    Set htmlDoc = CreateObject("MSHTML.HTMLFile")
    htmlDoc.Body.innerHTML = html
    StripHtmlTags = htmlDoc.Body.innerText
    Set htmlDoc = Nothing
    On Error GoTo 0
End Function

