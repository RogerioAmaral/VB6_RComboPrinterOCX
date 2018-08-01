VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RComboPrinter 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   ForwardFocus    =   -1  'True
   ScaleHeight     =   795
   ScaleWidth      =   5175
   ToolboxBitmap   =   "RComboPrinter.ctx":0000
   Begin MSComctlLib.ImageCombo cmbImpressoras 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      ImageList       =   "imlImpressoras"
   End
   Begin MSComctlLib.ImageList imlImpressoras 
      Left            =   4110
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "RComboPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_ToolTipData = ""
Const m_def_CampoData = 0
'Const m_def_CampoData = 1
Const m_def_DataMember = ""
'Const m_def_Posição = ""
'Const m_def_Coluna = ""
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Text = ""
'Const m_def_MaxLength = 0
'Const m_def_Posição = 0
'Const m_def_Coluna = ""
'Const m_def_Nulo = 0
'Property Variables:
Dim m_ToolTipData As String
Dim m_CampoData As Boolean
'Dim m_CampoData As Boolean
Dim m_DataMember As String
'Dim m_Posição As String
'Dim m_Coluna As String
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_Text As String
'Dim m_MaxLength As Long
'Dim m_Posição As Integer
'Dim m_Coluna As String
'Dim m_Nulo As Boolean
'Event Declarations:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()



Private hIcon As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Dim bolSair As Boolean

Dim strConfigs As String


Public Function ListarImpressoras(bolMostarIcones As Boolean) As Integer


   Dim IID_IShellFolderL As IShellFolderEx_TLB.Guid
    Dim IID_IPictureL(0 To 3) As Long
    Dim pidlPrintersL()  As Byte
    Dim pidlCurrentL()   As Byte
    Dim pidlAbsoluteL()  As Byte
    Dim pDesktopFolderL  As IShellFolder
    Dim pPrintersFolderL As IShellFolder
    Dim pEnumIdsL        As IEnumIDList
    Dim lPtrL            As Long
    Dim uInfoL           As SHFILEINFO
    Dim uPictL           As PictDesc
    Dim sPrinterNameL    As String
    Dim oPrinterIconL    As StdPicture
    Dim strImpDefaultL   As String

    '--- init consts
    IID_IShellFolderL.Data1 = &H214E6 '--- {000214E6-0000-0000-C000-000000000046}
    IID_IShellFolderL.Data4(0) = &HC0
    IID_IShellFolderL.Data4(7) = &H46
    IID_IPictureL(0) = &H7BF80980 '--- {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    IID_IPictureL(1) = &H101ABF32
    IID_IPictureL(2) = &HAA00BB8B
    IID_IPictureL(3) = &HAB0C3000
    '--- init local vars
    uPictL.Size = Len(uPictL)
    uPictL.Type = vbPicTypeIcon
    Call SHGetDesktopFolder(pDesktopFolderL)
    '--- retrieve enumerator of Printers virtual folder
    Call SHGetSpecialFolderLocation(0, CSIDL_PRINTERS, lPtrL)
    pidlPrintersL = pvToPidl(lPtrL)
    Call pDesktopFolderL.BindToObject(VarPtr(pidlPrintersL(0)), 0, IID_IShellFolderL, pPrintersFolderL)
    Call pPrintersFolderL.EnumObjects(0, SHCONTF_NONFOLDERS, pEnumIdsL)
    '--- loop printers
    Dim iIndexL As Integer
    iIndexL = 0
    
    ListarImpressoras = 0
        
    
    strImpDefaultL = Printer.DeviceName

    
    cmbImpressoras.ComboItems.Clear
    imlImpressoras.ListImages.Clear
    
    Do While pEnumIdsL.Next(1, lPtrL, 0) = 0 '--- S_OK
        pidlCurrentL = pvToPidl(lPtrL)
        '--- combine pidls: Printers + Current
        iIndexL = iIndexL + 1
        
        ReDim pidlAbsoluteL(0 To UBound(pidlPrintersL) + UBound(pidlCurrentL))
        Call CopyMemory(pidlAbsoluteL(0), pidlPrintersL(0), UBound(pidlPrintersL) - 1)
        Call CopyMemory(pidlAbsoluteL(UBound(pidlPrintersL) - 1), pidlCurrentL(0), UBound(pidlCurrentL) - 1)
        '--- retrieve info
        Call SHGetFileInfo(pidlAbsoluteL(0), 0, uInfoL, Len(uInfoL), SHGFI_PIDL Or SHGFI_DISPLAYNAME Or SHGFI_ICON)
        sPrinterNameL = Left(uInfoL.szDisplayName, InStr(uInfoL.szDisplayName, Chr$(0)) - 1)
        '--- extract icon
        If bolMostarIcones = True Then
            uPictL.hBmpOrIcon = uInfoL.hIcon
            Call OleCreatePictureIndirect(uPictL, IID_IPictureL(0), True, oPrinterIconL)
            '--- show
            imlImpressoras.ListImages.Add , , oPrinterIconL
            cmbImpressoras.ComboItems.Add , , sPrinterNameL, iIndexL, iIndexL
        Else
            cmbImpressoras.ComboItems.Add , , sPrinterNameL
        End If
        
        
        If strImpDefaultL = sPrinterNameL Then
            ListarImpressoras = iIndexL
        End If
        
        'Me.Print sPrinterName
        'MsgBox sPrinterName
    Loop

If ListarImpressoras = 0 Then ListarImpressoras = 1


Set cmbImpressoras.SelectedItem = cmbImpressoras.ComboItems(ListarImpressoras)

Call UserControl_Resize

End Function

Private Function pvToPidl(ByVal lPtr As Long) As Byte()
    Dim lTotal      As Long
    Dim nSize       As Integer
    Dim baPidl()    As Byte

    Do
        Call CopyMemory(nSize, ByVal (lPtr + lTotal), 2)
        lTotal = lTotal + nSize
    Loop While nSize <> 0
    ReDim baPidl(0 To lTotal + 1)
    Call CopyMemory(baPidl(0), ByVal lPtr, lTotal + 2)
    Call CoTaskMemFree(lPtr)
    pvToPidl = baPidl
End Function






Private Sub cmbImpressoras_Change()
    RaiseEvent Change
End Sub

Private Sub cmbImpressoras_Click()
Dim strImpressoraSelecionadaL As String

On Error Resume Next

strImpressoraSelecionadaL = cmbImpressoras.SelectedItem.Text


Dim p As Printer

For Each p In Printers
    If UCase(p.DeviceName) = UCase(strImpressoraSelecionadaL) Then
        Set Printer = p
        Exit For
    End If
Next p

'Printer.PaperSize = vbPRPSA4
'Printer.Orientation = vbPRORPortrait

Dim W As New WshNetwork
W.SetDefaultPrinter (strImpressoraSelecionadaL)
Set W = Nothing

Dim ppRet As New ClassPrinter

'ppRet.SetPaperSize Printer.DeviceName, 9
'ppRet.SetOrientation Printer.DeviceName, 1

'Printer.PaperSize = vbPRPSA4
'Printer.Orientation = vbPRORPortrait

Dim cpPrinter As New ClassPrinter
Dim lRetPaper As Long
Dim lRetOrient As Long

lRetPaper = cpPrinter.GetPaperSize(Printer.DeviceName)
lRetOrient = cpPrinter.GetOrientation(Printer.DeviceName)
    
    Dim strPapel As String
    Dim strOrientacao As String
    
    strPapel = RetPaperSize(lRetPaper)
    strOrientacao = IIf(lRetOrient = 1, "Retrato", "Paisagem")
    
    If lRetPaper <> Printer.PaperSize Then
        Printer.PaperSize = lRetPaper
        strPapel = strPapel & " (" & Printer.PaperSize & ")"
    End If
    
    If lRetOrient <> Printer.Orientation Then
        strOrientacao = strOrientacao & "/" & IIf(Printer.Orientation = 1, "Retrato", "Paisagem")
    End If
    
    cmbImpressoras.ToolTipText = Printer.DeviceName & " | Papel: " & strPapel & " (" & lRetPaper & "-" & Printer.PaperSize & ") | Orientação: " & strOrientacao & " (" & lRetOrient & "-" & Printer.Orientation & ")"
    
    strConfigs = cmbImpressoras.ToolTipText
    
RaiseEvent Click

End Sub

Private Sub cmbImpressoras_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_InitProperties()
   m_ForeColor = m_def_ForeColor
   m_Enabled = m_def_Enabled
   Set m_Font = Ambient.Font
   m_BackStyle = m_def_BackStyle
   m_BorderStyle = m_def_BorderStyle
   m_Text = m_def_Text
'   m_MaxLength = m_def_MaxLength
'   m_Posição = m_def_Posição
'   m_Coluna = m_def_Coluna
'   m_Nulo = m_def_Nulo
'   m_Posição = m_def_Posição
'   m_Coluna = m_def_Coluna
    m_DataMember = m_def_DataMember
'    m_CampoData = m_def_CampoData
    m_CampoData = m_def_CampoData
    m_ToolTipData = m_def_ToolTipData
        

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


   m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
   m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
   Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
   m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_Text = PropBag.ReadProperty("Text", m_def_Text)
'   m_MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
'   m_Posição = PropBag.ReadProperty("Posição", m_def_Posição)
'   m_Coluna = PropBag.ReadProperty("Coluna", m_def_Coluna)
'   m_Nulo = PropBag.ReadProperty("Nulo", m_def_Nulo)
'   Text1.Text = PropBag.ReadProperty("Text", "")
'   Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
'   Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
'   m_Posição = PropBag.ReadProperty("Posição", m_def_Posição)
'   m_Coluna = PropBag.ReadProperty("Coluna", m_def_Coluna)
    m_DataMember = PropBag.ReadProperty("DataMember", m_def_DataMember)
'    m_CampoData = PropBag.ReadProperty("CampoData", m_def_CampoData)
    m_CampoData = PropBag.ReadProperty("CampoData", m_def_CampoData)
    
    m_ToolTipData = PropBag.ReadProperty("ToolTipData", m_def_ToolTipData)
    'Text1.ToolTipText = m_ToolTipData
    cmbImpressoras.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
'    Label1.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)

End Sub

Private Sub UserControl_Resize()
    
    cmbImpressoras.Width = UserControl.Width
    
    If cmbImpressoras.Height = 570 Then
        UserControl.Height = 570
    Else
        UserControl.Height = 375
        cmbImpressoras.Height = 375
    End If
    
    
End Sub


Public Function RetornarConfig() As String
    RetornarConfig = strConfigs
End Function


Public Property Get ForeColor() As Long
   ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
   m_ForeColor = New_ForeColor
   PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
   Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set m_Font = New_Font
   PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
   BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
   m_BackStyle = New_BackStyle
   PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
   BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
   m_BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
   Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
   Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
   Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
'   Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
'   Call PropBag.WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
'   Call PropBag.WriteProperty("Posição", m_Posição, m_def_Posição)
'   Call PropBag.WriteProperty("Coluna", m_Coluna, m_def_Coluna)
'   Call PropBag.WriteProperty("Nulo", m_Nulo, m_def_Nulo)
'   Call PropBag.WriteProperty("Text", Text1.Text, "")
'   Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
'   Call PropBag.WriteProperty("Caption", Label1.Caption, "Label1")
'   Call PropBag.WriteProperty("Posição", m_Posição, m_def_Posição)
'   Call PropBag.WriteProperty("Coluna", m_Coluna, m_def_Coluna)
    Call PropBag.WriteProperty("DataMember", m_DataMember, m_def_DataMember)
    Call PropBag.WriteProperty("Locked", cmbImpressoras.Locked, False)
'    Call PropBag.WriteProperty("CampoData", m_CampoData, m_def_CampoData)
'    Call PropBag.WriteProperty("CampoData", m_CampoData, m_def_CampoData)
    Call PropBag.WriteProperty("ToolTipData", m_ToolTipData, m_def_ToolTipData)
    Call PropBag.WriteProperty("BackColor", cmbImpressoras.BackColor, &H80000005)
'    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &HFFFFFF)

End Sub







Public Property Get Text() As String
   Text = cmbImpressoras.Text
End Property

Public Property Let Text(ByVal New_Text As String)
   cmbImpressoras.Text = New_Text
   PropertyChanged "Text"
   
End Property

Public Property Get DataMember() As String
    DataMember = m_DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    m_DataMember = New_DataMember
    PropertyChanged "DataMember"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Locked
Public Property Get Locked() As Boolean
    Locked = cmbImpressoras.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    cmbImpressoras.Locked = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = cmbImpressoras.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    cmbImpressoras.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property






