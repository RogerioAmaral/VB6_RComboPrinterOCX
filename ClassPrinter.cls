VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Type PRINTER_DEFAULTS
   pDatatype As Long
   pDevmode As Long
   DesiredAccess As Long
End Type
 
Private Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevmode As Long               ' Pointer to DEVMODE
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long    ' Pointer to SECURITY_DESCRIPTOR
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type
 
Private Type DEVMODE
   dmDeviceName As String * 32
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * 32
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
   dmICMMethod As Long
   dmICMIntent As Long
   dmMediaType As Long
   dmDitherType As Long
   dmReserved1 As Long
   dmReserved2 As Long
End Type
 
Private Const DM_ORIENTATION = &H1
Private Const DM_PAPERSIZE = &H2
Private Const DM_PAPERLENGTH = &H4
Private Const DM_PAPERWIDTH = &H8
Private Const DM_DEFAULTSOURCE = &H200
Private Const DM_PRINTQUALITY = &H400
Private Const DM_COLOR = &H800
Private Const DM_DUPLEX = &H1000
 
Private Const DM_IN_BUFFER = 8
Private Const DM_OUT_BUFFER = 2
 
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000          ' Allowed to read device information
Private Const WRITE_DAC = &H40000             ' Allowed to write device access control info
Private Const WRITE_OWNER = &H80000           ' Allowed to change the object owner
' Combining these for full access to a device  (DELETE + READ_CONTROL + WRITE_DAC + WRITE_OWNER):
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
 
Private Const SERVER_ACCESS_ADMINISTER = &H1  ' Access rights to administer print servers.
Private Const SERVER_ACCESS_ENUMERATE = &H2   ' Access rights to enumerate print servers.
Private Const PRINTER_ACCESS_ADMINISTER = &H4 ' Access rights for printers to perform administrative tasks.
Private Const PRINTER_ACCESS_USE = &H8        ' Access rights for printers for general use (printing, querying).
 
' Access which allows you to set duplex on or off
Private Const PRINTER_NORMAL_ACCESS = (READ_CONTROL Or PRINTER_ACCESS_USE)
 
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_LOCAL = &H2
 
Private Declare Function ClosePrinter Lib "winspool.drv" _
      (ByVal hPrinter As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" _
      Alias "DocumentPropertiesA" (ByVal hwnd As Long, _
      ByVal hPrinter As Long, ByVal pDeviceName As String, _
      ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, _
      ByVal fMode As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias _
      "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
      pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
      "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
      pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias _
      "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
      pPrinter As Byte, ByVal Command As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" _
      Alias "EnumPrintersA" _
      (ByVal flags As Long, ByVal name As String, ByVal Level As Long, _
      pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, _
      pcReturned As Long) As Long
 
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" _
      (ByVal RetVal As String, ByVal Ptr As Long) As Long
 
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenA" _
      (ByVal Ptr As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (pDest As Any, pSource As Any, ByVal cbLength As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Private Declare Function DeviceCapabilities Lib "winspool.drv" _
      Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
      ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
      ByVal dev As Long) As Long
 
 
Public Sub SetOrientation(ByVal sPrinterName As String, iOrientation As Long)
   SetPrinterProperty sPrinterName, DM_ORIENTATION, iOrientation
End Sub

Public Function GetOrientation(ByVal sPrinterName As String) As Long
  GetOrientation = GetPrinterProperty(sPrinterName, DM_ORIENTATION)
End Function
 
Public Sub SetPaperSize(ByVal sPrinterName As String, iPaperSize As Long)
   SetPrinterProperty sPrinterName, DM_PAPERSIZE, iPaperSize
End Sub
 
Public Function GetPaperSize(ByVal sPrinterName As String) As Long
  GetPaperSize = GetPrinterProperty(sPrinterName, DM_PAPERSIZE)
End Function
 
 
Public Sub SetColorMode(ByVal sPrinterName As String, iColorMode As Long)
   SetPrinterProperty sPrinterName, DM_COLOR, iColorMode
End Sub
 
Public Function GetColorMode(ByVal sPrinterName As String) As Long
  GetColorMode = GetPrinterProperty(sPrinterName, DM_COLOR)
End Function
 
Public Sub SetDuplex(ByVal sPrinterName As String, iDuplex As Long)
   SetPrinterProperty sPrinterName, DM_DUPLEX, iDuplex
End Sub
 
Public Function GetDuplex(ByVal sPrinterName As String) As Long
   GetDuplex = GetPrinterProperty(sPrinterName, DM_DUPLEX)
End Function
 
Public Sub SetPrintQuality(ByVal sPrinterName As String, iQuality As Long)
   SetPrinterProperty sPrinterName, DM_PRINTQUALITY, iQuality
End Sub
 
Public Function GetPrintQuality(ByVal sPrinterName As String) As Long
   GetPrintQuality = GetPrinterProperty(sPrinterName, DM_PRINTQUALITY)
End Function
 
Private Function SetPrinterProperty(ByVal sPrinterName As String, ByVal iPropertyType As Long, _
      ByVal iPropertyValue As Long) As Boolean
 
   'Code adapted from Microsoft KB article Q230743
 
    Dim hPrinter As Long          'handle for the current printer
    Dim pd As PRINTER_DEFAULTS
    Dim pinfo As PRINTER_INFO_2
    Dim dm As DEVMODE
 
    Dim yDevModeData() As Byte        'Byte array to hold contents
                                      'of DEVMODE structure
    Dim yPInfoMemory() As Byte        'Byte array to hold contents
                                      'of PRINTER_INFO_2 structure
    Dim iBytesNeeded As Long
    Dim iRet As Long
    Dim iJunk As Long
    Dim iCount As Long
       
    On Error GoTo cleanup
 
    pd.DesiredAccess = PRINTER_NORMAL_ACCESS
    iRet = OpenPrinter(sPrinterName, hPrinter, pd)
    If (iRet = 0) Or (hPrinter = 0) Then
       'Can't access current printer. Bail out doing nothing
       Exit Function
    End If
 
    'Get the size of the DEVMODE structure to be loaded
    iRet = DocumentProperties(0, hPrinter, sPrinterName, 0, 0, 0)
    If (iRet < 0) Then
       'Can't access printer properties.
       GoTo cleanup
    End If
 
    'Make sure the byte array is large enough
    'Some printer drivers lie about the size of the DEVMODE structure they
    'return, so an extra 100 bytes is provided just in case!
    ReDim yDevModeData(0 To iRet + 100) As Byte
       
    'Load the byte array
    iRet = DocumentProperties(0, hPrinter, sPrinterName, _
                VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
    If (iRet < 0) Then
       GoTo cleanup
    End If
 
    'Copy the byte array into a structure so it can be manipulated
    Call CopyMemory(dm, yDevModeData(0), Len(dm))
 
    If dm.dmFields And iPropertyType = 0 Then
       'Wanted property not available. Bail out.
       GoTo cleanup
    End If
 
    'Set the property to the appropriate value
    Select Case iPropertyType
    Case DM_ORIENTATION
       dm.dmOrientation = iPropertyValue
    Case DM_PAPERSIZE
       dm.dmPaperSize = iPropertyValue
    Case DM_PAPERLENGTH
       dm.dmPaperLength = iPropertyValue
    Case DM_PAPERWIDTH
       dm.dmPaperWidth = iPropertyValue
    Case DM_DEFAULTSOURCE
       dm.dmDefaultSource = iPropertyValue
    Case DM_PRINTQUALITY
       dm.dmPrintQuality = iPropertyValue
    Case DM_COLOR
       dm.dmColor = iPropertyValue
    Case DM_DUPLEX
       dm.dmDuplex = iPropertyValue
    End Select
       
    'Load the structure back into the byte array
    Call CopyMemory(yDevModeData(0), dm, Len(dm))
 
    'Tell the printer about the new property
    iRet = DocumentProperties(0, hPrinter, sPrinterName, _
          VarPtr(yDevModeData(0)), VarPtr(yDevModeData(0)), _
          DM_IN_BUFFER Or DM_OUT_BUFFER)
 
    If (iRet < 0) Then
       GoTo cleanup
    End If
 
    'The code above *ought* to be sufficient to set the property
    'correctly. Unfortunately some brands of Postscript printer don't
    'seem to respond correctly. The following code is used to make
    'sure they also respond correctly.
    Call GetPrinter(hPrinter, 2, 0, 0, iBytesNeeded)
    If (iBytesNeeded = 0) Then
       'Couldn't access shared printer settings
       GoTo cleanup
    End If
       
    'Set byte array large enough for PRINTER_INFO_2 structure
    ReDim yPInfoMemory(0 To iBytesNeeded + 100) As Byte
 
    'Load the PRINTER_INFO_2 structure into byte array
    iRet = GetPrinter(hPrinter, 2, yPInfoMemory(0), iBytesNeeded, iJunk)
    If (iRet = 0) Then
       'Couldn't access shared printer settings
       GoTo cleanup
    End If
 
    'Copy byte array into the structured type
    Call CopyMemory(pinfo, yPInfoMemory(0), Len(pinfo))
 
    'Load the DEVMODE structure with byte array containing
    'the new property value
    pinfo.pDevmode = VarPtr(yDevModeData(0))
       
    'Set security descriptor to null
    pinfo.pSecurityDescriptor = 0
      
    'Copy the PRINTER_INFO_2 structure back into byte array
    Call CopyMemory(yPInfoMemory(0), pinfo, Len(pinfo))
 
    'Send the new details to the printer
    iRet = SetPrinter(hPrinter, 2, yPInfoMemory(0), 0)
 
    'Indicate whether it all worked or not!
    SetPrinterProperty = CBool(iRet)
 
cleanup:
   'Release the printer handle
   If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)
       
   'Flush the message queue. If you don't do this,
   'you can get page fault errors when you try to
   'print a document immediately after setting a printer property.
   For iCount = 1 To 20
      DoEvents
   Next iCount
   End Function
 
Private Function GetPrinterProperty(ByVal sPrinterName As String, ByVal iPropertyType As Long) As Long
 
  'Code adapted from Microsoft KB article Q230743
 
  Dim hPrinter As Long
  Dim pd As PRINTER_DEFAULTS
  Dim dm As DEVMODE
 
  Dim yDevModeData() As Byte
  Dim iRet As Long
       
  On Error GoTo cleanup
       
  pd.DesiredAccess = PRINTER_NORMAL_ACCESS
       
  'Get the printer handle
  iRet = OpenPrinter(sPrinterName, hPrinter, pd)
  If (iRet = 0) Or (hPrinter = 0) Then
     'Couldn't access the printer
      Exit Function
  End If
 
  'Find out how many bytes needed for the printer properties
  iRet = DocumentProperties(0, hPrinter, sPrinterName, 0, 0, 0)
  If (iRet < 0) Then
     'Couldn't access printer properties
      GoTo cleanup
  End If
 
  'Make sure the byte array is large enough, including the
  '100 bytes extra in case the printer driver is lying.
  ReDim yDevModeData(0 To iRet + 100) As Byte
       
  'Load the printer properties into the byte array
  iRet = DocumentProperties(0, hPrinter, sPrinterName, _
              VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
  If (iRet < 0) Then
     'Couldn't access printer properties
     GoTo cleanup
  End If
 
 
 
  'Copy the byte array to the DEVMODE structure
  Call CopyMemory(dm, yDevModeData(0), Len(dm))
 
  If Not dm.dmFields And iPropertyType = 0 Then
     'Requested property not available on this printer.
     GoTo cleanup
  End If
 
  'Get the value of the requested property
  Select Case iPropertyType
  Case DM_ORIENTATION
     GetPrinterProperty = dm.dmOrientation
  Case DM_PAPERSIZE
     GetPrinterProperty = dm.dmPaperSize
  Case DM_PAPERLENGTH
     GetPrinterProperty = dm.dmPaperLength
  Case DM_PAPERWIDTH
     GetPrinterProperty = dm.dmPaperWidth
  Case DM_DEFAULTSOURCE
     GetPrinterProperty = dm.dmDefaultSource
  Case DM_PRINTQUALITY
     GetPrinterProperty = dm.dmPrintQuality
  Case DM_COLOR
     GetPrinterProperty = dm.dmColor
  Case DM_DUPLEX
     GetPrinterProperty = dm.dmDuplex
  End Select
       
cleanup:
   'Release the printer handle
   If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)
 
End Function

