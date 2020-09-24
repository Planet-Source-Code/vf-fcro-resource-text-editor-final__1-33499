VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
exitX = True
Unload Me
End Sub

Private Sub Form_Load()
Caption = F2CAP
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Const SC_CLOSE = &HF060
Dim hMenu As Long
hMenu = GetSystemMenu(hWnd, 0&)
If hMenu Then
Call DeleteMenu(hMenu, SC_CLOSE, 0)
DrawMenuBar (hWnd)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not IsNumeric(Text1) Then Text1 = "": Exit Sub
If CLng(Text1) < 0 Or CLng(Text1) > 65535 Then Text1 = "": Exit Sub
IDDP = LongToInt(Text1)
exitX = False
Unload Me
End If
End Sub
