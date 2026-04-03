# VBA-FileParser

A comprehensive VBA class library for extracting text from various file formats including PDF, Excel, Word, TXT, HTML, XML, Images, and Email files (.msg, .eml).

## Features

- **PDF**: Extract text using PdfParser (supports PdfTXT, OCR, and Word COM fallback)
- **Excel**: Extract text from .xls, .xlsx, .xlsb, .xlsm files (with PDF/OCR fallback)
- **Word**: Extract text from .doc, .docx, .docm files (with PDF/OCR fallback)
- **Text**: Plain text and CSV files
- **HTML**: Extract text from HTML files
- **XML**: Extract text from XML files
- **Images**: OCR extraction from PNG, JPG, JPEG, GIF, BMP, WebP, TIFF, TIF
- **Email**: 
  - `.msg` files (using Outlook COM)
  - `.eml` files (parsed from raw file)
  - Recursive attachment extraction

## Installation

1. Clone the repository with submodules:
   ```bash
   git clone --recurse-submodules https://github.com/your-repo/VBA-FileParser.git
   ```
   Or initialize submodules after cloning:
   ```bash
   git submodule update --init --recursive
   ```

2. In your VBA project:
   - Add `FileParser.cls` to your project
   - Ensure the PdfParser submodule is accessible (add reference to helpers/PdfParser/PdfParser.cls)

3. Required References (in VBA Editor > Tools > References):
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

' For email: item 1 = body, items 2+ = attachment text
Set pages = fp.ExtractPages("C:\path\to\email.msg")
```

### Check File Type

```vba
Dim fp As New FileParser
fp.ExtractText "C:\path\to\document.pdf"

Debug.Print fp.LastType  ' Output: "Pdf"
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
        Debug.Print "File: " & file & " Type: " & fp.LastType
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
- **Returns**: Collection where each item is text from a page/sheet/item
  - For PDF: Each page
  - For Excel: Each sheet
  - For Word: Each page
  - For Email: Item 1 = body, items 2+ = attachment text

#### LastType() As String
- **Returns**: The type of extractor that succeeded (e.g., "Pdf", "Excel", "Word", "Msg", "Eml", "Text", "Html", "Xml", "Image", "Unknown")

#### Parser As PdfParser (Property)
- **Returns**: Reference to the underlying PdfParser instance for tuning

## Constants

- `MaxRecursionDepth`: 5 (prevents infinite loops when processing nested email attachments)

## File Structure

```
FileParser/
├── FileParser.cls          # Main class
├── helpers/
│   └── PdfParser/          # Submodule for PDF extraction
├── README.md
└── LICENSE
```

## License

MIT License - See [LICENSE](LICENSE) file for details.