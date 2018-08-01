Attribute VB_Name = "modComboPrinter"
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'Icon Sizes in pixels
Public Const LARGE_ICON As Integer = 32
Public Const SMALL_ICON As Integer = 16
Public Const MAX_PATH = 260

Public Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Public Const SHGFI_LARGEICON = &H0       'Large icon
Public Const SHGFI_SMALLICON = &H1       'Small icon
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400

Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


Public Const CSIDL_PRINTERS    As Long = &H4
Public Const SHGFI_PIDL        As Long = &H8
Public Const SHGFI_ICON        As Long = &H100

Public Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As IShellFolder) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFOO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, ppRet As IPicture) As Long


Public Type PictDesc
    Size                As Long
    Type                As Long
    hBmpOrIcon          As Long
    hPal                As Long
End Type



'Public Function ShowPrinterProperties(ByVal DeviceName As String, _
'    ByVal ParentHWnd As Long) As Boolean
'    Dim PrinterDef As PRINTER_DEFAULTS
'    Dim hPrinter As Long
'    Const PRINTER_ALL_ACCESS = &HF000C
'
'    PrinterDef.DesiredAccess = PRINTER_ALL_ACCESS
'
'    If OpenPrinter(DeviceName, hPrinter, PrinterDef) Then
'        ShowPrinterProperties = PrinterProperties(ParentHWnd, hPrinter)
'        ClosePrinter hPrinter
'    End If
'End Function


Public Function RetPaperSize(intPapel As Long) As String

Select Case intPapel
    Case 1
        'vbPRPSLetter
        RetPaperSize = "Carta/Letter, 8 1/2 x 11 in."
        'vbPRPSLetterSmall
    Case 2
        RetPaperSize = "Carta/Letter Small, 8 1/2 x 11 in."
        'vbPRPSTabloid
    Case 3
        RetPaperSize = "Tabloid, 11 x 17 in."
        'vbPRPSLedger
    Case 4
        RetPaperSize = "Ledger, 17 x 11 in."
        'vbPRPSLegal
    Case 5
        RetPaperSize = "Legal, 8 1/2 x 14 in."
        'vbPRPSStatement
    Case 6
        RetPaperSize = "Statement, 5 1/2 x 8 1/2 in."
        'vbPRPSExecutive
    Case 7
        RetPaperSize = "Executive, 7 1/2 x 10 1/2 in."
        'vbPRPSA3
    Case 8
        RetPaperSize = "A3, 297 x 420 mm"
        'vbPRPSA4
    Case 9
        RetPaperSize = "A4, 210 x 297 mm"
        'vbPRPSA4Small
    Case 10
        RetPaperSize = "A4 Small, 210 x 297 mm"
        'vbPRPSA5
    Case 11
        RetPaperSize = "A5, 148 x 210 mm"
        'vbPRPSB4
    Case 12
        RetPaperSize = "B4, 250 x 354 mm"
        'vbPRPSB5
    Case 13
        RetPaperSize = "B5, 182 x 257 mm"
        'vbPRPSFolio
    Case 14
        RetPaperSize = "Folio, 8 1/2 x 13 in."
        'vbPRPSQuarto
    Case 15
        RetPaperSize = "Quarto, 215 x 275 mm"
        'vbPRPS10x14
    Case 16
        RetPaperSize = "10 x 14 in."
        'vbPRPS11x17
    Case 17
        RetPaperSize = "11 x 17 in."
        'vbPRPSNote
    Case 18
        RetPaperSize = "Note, 8 1/2 x 11 in."
        'vbPRPSEnv9
    Case 19
        RetPaperSize = "Envelope #9, 3 7/8 x 8 7/8 in."
        'vbPRPSEnv10
    Case 20
        RetPaperSize = "Envelope #10, 4 1/8 x 9 1/2 in."
        'vbPRPSEnv11
    Case 21
        RetPaperSize = "Envelope #11, 4 1/2 x 10 3/8 in."
        'vbPRPSEnv12
    Case 22
        RetPaperSize = "Envelope #12, 4 1/2 x 11 in."
        'vbPRPSEnv14
    Case 23
        RetPaperSize = "Envelope #14, 5 x 11 1/2 in."
        'vbPRPSCSheet
    Case 24
        RetPaperSize = "C size sheet"
        'vbPRPSDSheet
    Case 25
        RetPaperSize = "D size sheet"
        'vbPRPSESheet
    Case 26
        RetPaperSize = "E size sheet"
        'vbPRPSEnvDL
    Case 27
        RetPaperSize = "Envelope DL, 110 x 220 mm"
        'vbPRPSEnvC3
    Case 29
        RetPaperSize = "Envelope C3, 324 x 458 mm"
        'vbPRPSEnvC4
    Case 30
        RetPaperSize = "Envelope C4, 229 x 324 mm"
        'vbPRPSEnvC5
    Case 28
        RetPaperSize = "Envelope C5, 162 x 229 mm"
        'vbPRPSEnvC6
    Case 31
        RetPaperSize = "Envelope C6, 114 x 162 mm"
        'vbPRPSEnvC65
    Case 32
        RetPaperSize = "Envelope C65, 114 x 229 mm"
        'vbPRPSEnvB
    Case 33
        RetPaperSize = "Envelope B4, 250 x 353 mm"
        'vbPRPSEnvB5
    Case 34
        RetPaperSize = "Envelope B5, 176 x 250 mm"
        'vbPRPSEnvB6
    Case 35
        RetPaperSize = "Envelope B6, 176 x 125 mm"
        'vbPRPSEnvItaly
    Case 36
        RetPaperSize = "Envelope, 110 x 230 mm"
        'vbPRPSEnvMonarch
    Case 37
        RetPaperSize = "Envelope Monarch, 3 7/8 x 7 1/2 in."
        'vbPRPSEnvPersonal
    Case 38
        RetPaperSize = "Envelope, 3 5/8 x 6 1/2 in."
        'vbPRPSFanfoldUS
    Case 39
        RetPaperSize = "U.S. Standard Fanfold, 14 7/8 x 11 in."
        'vbPRPSFanfoldStdGerman
    Case 40
        RetPaperSize = "German Standard Fanfold, 8 1/2 x 12 in."
        'vbPRPSFanfoldLglGerman
    Case 41
        RetPaperSize = "German Legal Fanfold, 8 1/2 x 13 in."
        'vbPRPSUser
    Case 256
        RetPaperSize = "User Defined"
    Case Else
        RetPaperSize = "Unknown"
End Select

End Function


