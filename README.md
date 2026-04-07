# VBA-FileParser

A comprehensive VBA class library for extracting text from various file formats including PDF (using [PdfParser](https://github.com/rafael-yml/VBA-PdfParser)), Excel, Word, TXT, HTML, XML, Images, Email files (.msg, .eml), and ZIP archives.

## Features

- **PDF**: Extract text using PdfParser (supports PdfTXT, OCR, and Word COM fallback)
- **Excel**: Extract text from .xls, .xlsx, .xlsb, .xlsm files (with PDF/OCR fallback)
- **Word**: Extract text from .doc, .docx, .docm files (with PDF/OCR fallback)
- **Text**: Plain text and CSV files
- **HTML**: Extract text from HTML files
- **XML**: Extract text from XML files
- **Images**: OCR extraction from PNG, JPG, JPEG, GIF, BMP, WebP, TIFF, TIF (using Windows.Media.Ocr)
- **ZIP**: Extract text from .zip files, including nested ZIPs (using Shell.Application)
- **Email**: 
  - `.msg` files (using Outlook COM)
  - `.eml` files (parsed from raw file, supports quoted-printable and Base64)
  - Recursive attachment extraction with depth limit
- **Outlook caching**: Reuses Outlook instance for better performance

## Installation

1. In your VBA project, import the following files:
   - `FileParser.cls` (main class)
   - From `PdfParser/` (submodule):
     - `PdfParser.cls` (pdf wrapper class)
     - `PdfTXT/PdfTXT.cls`
     - `PdfWRT/PdfWRT.cls`
     - `WinOCR/WinOCR.cls`
     - `WdCOM/WdCOM.cls`

2. Required References (in VBA Editor > Tools > References):
   - Microsoft Scripting Runtime
   - Microsoft XML, v6.0 (for XML parsing)
   - Microsoft HTML Object Library (for HTML parsing)
   - Microsoft Outlook XX Object Library (for .msg files)
   - Microsoft Word XX Object Library (for Word files)
   - Microsoft Excel XX Object Library (for Excel files)

## Usage

### Basic Extraction

```vba
Dim fp As New FileParser
Dim text As String

' Extract all text from a file
text = fp.ExtractText("C:\path\to\document.pdf")
Debug.Print text

' Extract without attachments (for email)
text = fp.ExtractText("C:\path\to\email.msg", extractAttachments:=False)
```

### Per-Page/Per-Item Extraction

```vba
Dim fp As New FileParser
Dim pages As Collection
Dim i As Long

Set pages = fp.ExtractPages("C:\path\to\document.pdf")

For i = 1 To pages.Count
    Debug.Print "Page " & i & ": " & pages(i)
Next i

' For email: item 1 = subject, item 2 = body, items 3+ = attachment text
Set pages = fp.ExtractPages("C:\path\to\email.msg")
```

### Check File Type and Status

```vba
Dim fp As New FileParser
Dim text As String

text = fp.ExtractText("C:\path\to\document.pdf")

Debug.Print fp.LastType    ' Output: "Pdf"
Debug.Print fp.LastStatus  ' Output: 0 (success), or error code
```

### Access PdfParser for Tuning

```vba
Dim fp As New FileParser

' Access underlying PdfParser
With fp.Parser
    ' Configure PdfParser options here
    ' e.g., .TXT property for text extraction settings
End With
```

## Example: Processing Multiple Files

```vba
Sub ProcessFiles()
    Dim fp As New FileParser
    Dim files As Collection
    Dim file As Variant
    Dim text As String
    
    Set files = New Collection
    files.Add "C:\docs\report.pdf"
    files.Add "C:\docs\data.xlsx"
    files.Add "C:\docs\email.msg"
    
    For Each file In files
        text = fp.ExtractText(CStr(file))
        Debug.Print "File: " & file & " Type: " & fp.LastType & " Status: " & fp.LastStatus
        Debug.Print "Text Length: " & Len(text)
    Next
End Sub
```

## API Reference

### Methods

#### ExtractText(filePath As String, Optional extractAttachments As Boolean = True) As String
- **filePath**: Full path to the file to extract text from
- **extractAttachments**: For email files, whether to recursively extract attachment text (default: True)
- **Returns**: All extracted text concatenated with line breaks

#### ExtractPages(filePath As String) As Collection
- **filePath**: Full path to the file to extract pages from
- **extractAttachments**: For email files, whether to recursively extract attachment text (default: True)
- **Returns**: Collection where each item is text from a page/sheet/item
  - PDF: Each page
  - Excel: Each sheet
  - Word: Each page
  - Text/HTML/XML/Image: Single item with full content
  - Email: Item 1 = subject, Item 2 = body, Items 3+ = attachment text

#### LastType() As String
- **Returns**: The type of extractor that succeeded (e.g., "Pdf", "Excel", "Word", "Msg", "Eml", "Text", "Html", "Xml", "Image", "Unknown", "TooLarge")

#### LastStatus() As Long
- **Returns**: Status code of the last extraction
  - `0` = Success
  - `1` = File not found
  - `2` = File too large (>30MB)
  - `10` = Excel: failed to open
  - `11` = Excel: no content extracted
  - `20` = Word: failed to create instance
  - `21` = Word: failed to open document
  - `30` = Text: failed to open file
  - `31` = Text: no content
  - `40` = HTML: failed to open file
  - `41` = HTML: no content
  - `50` = XML: failed to parse
  - `51` = XML: no content
  - `60` = Image: OCR failed
  - `70` = Msg: Outlook not available
  - `71` = Msg: failed to open item
  - `72` = Msg: no content
  - `80` = Eml: failed to open file
  - `81` = Eml: no content
  - `90` = Zip: failed to open
  - `91` = Zip: no content extracted

#### Parser As PdfParser (Property)
- **Returns**: Reference to the underlying PdfParser instance for tuning

## Constants

- `MaxRecursionDepth`: 5 (prevents infinite loops when processing nested email attachments)
- `MaxFileSizeBytes`: 30MB (files larger than this return status 2)

## File Structure

```
FileParser/
├── FileParser.cls              # Main class
├── PdfParser/                  # Submodule (with nested submodules)
│   ├── PdfParser.cls
│   ├── PdfTXT/
│   ├── PdfWRT/
│   ├── WinOCR/
│   └── WdCOM/
├── README.md
└── LICENSE
```
---

## Dependencies

| Helper | What it does |
|---|---|
| [VBA-PdfParser](https://github.com/rafael-yml/VBA-PdfParser) | Wrapper function for fully-featured PDF text extraction (uses the classes below) |
| [VBA-PdfTXT](https://github.com/rafael-yml/VBA-PdfTXT) | Pure VBA PDF text extraction |
| [VBA-PdfWRT](https://github.com/rafael-yml/VBA-PdfWRT) | PDF → PNG via WinRT |
| [VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR) | Image → text via Windows OCR |
| [VBA-WdCOM](https://github.com/rafael-yml/VBA-WdCOM) | Word COM automation fallback |

All helpers are included as git submodules in the `PdfParser/` directory.

---

## License

MIT License. See [LICENSE](LICENSE) for details.

---

## Credits

- [VBA-PdfParser](https://github.com/rafael-yml/VBA-PdfParser)
- [VBA-PdfTXT](https://github.com/rafael-yml/VBA-PdfTXT)
- [VBA-PdfWRT](https://github.com/rafael-yml/VBA-PdfWRT)
- [VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR)
- [VBA-WdCOM](https://github.com/rafael-yml/VBA-WdCOM)

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)