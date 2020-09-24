VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Resource Text Editor V0.99 by Vanja Fuckar,Email: INGA@VIP.HR"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Height          =   615
      Left            =   240
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Open Resource File"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   615
      Left            =   11160
      Picture         =   "Form1.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Height          =   615
      Left            =   2040
      Picture         =   "Form1.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Change Entry Id"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Height          =   615
      Left            =   2640
      Picture         =   "Form1.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "New/Clear All"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   5640
      Picture         =   "Form1.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Information"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   5040
      Picture         =   "Form1.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Open Text File"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   1440
      Picture         =   "Form1.frx":3D86
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "View Entry"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   3240
      Picture         =   "Form1.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Delete Entry"
      Top             =   120
      Width           =   615
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7230
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   840
      Picture         =   "Form1.frx":4F1A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save Resource File"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   4440
      Picture         =   "Form1.frx":57E4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Add Entry"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00ACC2A5&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Other Resources Count:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C00000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CB8472&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Text Id"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F1D2C9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Text Length:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spath As String
Dim sFile As String

Private Sub Command1_Click()
If Len(Text1) = 0 Or Len(Text1) > 65535 Then Exit Sub
F2CAP = "Entry ID Range 0-65535"
Form2.Show 1
If exitX = True Then Exit Sub
PutData IntToLong(IDDP), Text1
End Sub

Private Sub Command10_Click()
aa = GetOpenFilePath(hWnd, "RES (*.res)" & vbNullChar & "*.res", 0, sFile, "", "Open Resource File", spath)
If aa = False Then Exit Sub
Open spath For Binary As #1
ReDim ResData(LOF(1) - 1)
Get #1, , ResData
Close #1
countX = 0
Erase data
Erase id
List1.Clear
EnumRESFile
InsertDataX
Set StringDataX = Nothing
Set StringIDX = Nothing
Label4 = "Other Resources Count:" & RestDataX.Count
End Sub




Private Sub Command2_Click()
If List1.ListCount = 0 And countX = 0 And RestDataX.Count = 0 Then Exit Sub
aa = GetSaveFilePath(hWnd, "RES (*.res)" & vbNullChar & "*.res", 0, "", "", "", "Save Resource File", spath)
If aa = False Then Exit Sub
If Dir(spath) <> "" Then Kill spath
SaveRES spath
End Sub

Private Sub Command3_Click()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then Exit Sub
DeleteEntry List1.ListIndex
End Sub

Private Sub Command4_Click()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then Exit Sub
Text1 = data(List1.ListIndex)
End Sub

Private Sub Command5_Click()
aa = GetOpenFilePath(hWnd, "", 0, sFile, "", "Open Any Text File", spath)
If aa = False Then Exit Sub
If FileLen(spath) > 65535 Or FileLen(spath) = 0 Then
MsgBox "File could not exceed 65535 Bytes of Length!", vbCritical, "Error"
Exit Sub
End If
Dim str1 As String
Open spath For Binary As #1
str1 = Space(LOF(1))
Get #1, , str1
Close #1
Text1 = str1
str1 = ""
End Sub

Private Sub Command6_Click()
Dim str1x As String
str1x = "Warning:Maximum Length per row:65535 Bytes!!! Do not try to put Overloaded Row,because that will occur an Error!" & vbCrLf
str1x = str1x & "VB Resource Editor will not correctly display Table which have Rows with length over 32767 bytes!" & vbCrLf
str1x = str1x & "That is not an Error,so dont be worried about incorrect table display!"
MsgBox str1x, vbOKOnly, "Info"
End Sub

Private Sub Command7_Click()
Erase data
Erase id
List1.Clear
countX = 0
Set RestDataX = Nothing
Set StringDataX = Nothing
Set StringIDX = Nothing
Label4 = "Other Resources Count: 0"
End Sub

Private Sub Command8_Click()
If List1.ListCount = 0 Or List1.ListIndex = -1 Then Exit Sub
F2CAP = "Entry ID Range 0-65535"
Form2.Show 1
If exitX Then Exit Sub
Dim entr As Long
entr = EntrySearch(id, IntToLong(IDDP))
If entr <> -1 And entr <> List1.ListIndex Then MsgBox "Entry ID duplicate detected!", vbInformation, "Information": Exit Sub
id(List1.ListIndex) = IntToLong(IDDP)
Dim Mstrlen As Long
Dim stringX As String
stringX = data(List1.ListIndex)
Mstrlen = 35
If Len(stringX) < 35 Then Mstrlen = Len(stringX)
LB.List(List1.ListIndex) = IntToLong(IDDP) & vbTab & Left(stringX, Mstrlen)
QuickSortMe id, 0, UBound(id)
End Sub

Private Sub Command9_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Shape1.Width = ScaleWidth
Label1.Width = ScaleWidth - Label2.Width - Label3.Width
Label4 = "Other Resources Count: 0"
Set LB = List1
Dim tabs(1) As Long
tabs(0) = 50
tabs(1) = 0
SendMessage List1.hWnd, LB_SETTABSTOPS, 1, tabs(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase data
Erase id
Set RestDataX = Nothing
Set StringDataX = Nothing
Set StringIDX = Nothing
End Sub

Private Sub Text1_Change()
Label1 = "Text Length:" & Len(Text1)
End Sub
