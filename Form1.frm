VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   330
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1995
      Width           =   8160
   End
   Begin Project1.RComboPrinter RComboPrinter1 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1350
      Width           =   6090
      _extentx        =   10742
      _extenty        =   661
      font            =   "Form1.frx":0000
      locked          =   -1  'True
      backcolor       =   16777215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   1470
      TabIndex        =   0
      Top             =   180
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    RComboPrinter1.ListarImpressoras True
    
    
End Sub


Private Sub RComboPrinter1_Change()
    Debug.Print "RComboPrinter1_Change"
    Debug.Print RComboPrinter1.Text
End Sub

Private Sub RComboPrinter1_Click()
    Debug.Print "RComboPrinter1_Click"
    Debug.Print RComboPrinter1.Text
End Sub

Private Sub RComboPrinter1_DblClick()
Debug.Print "RComboPrinter1_DblClick"
End Sub

Private Sub RComboPrinter1_KeyPress(KeyAscii As Integer)
Debug.Print "RComboPrinter1_KeyPress:" & KeyAscii

End Sub

Private Sub Text1_Click()
        Text1.Text = RComboPrinter1.RetornarConfig
        
End Sub
