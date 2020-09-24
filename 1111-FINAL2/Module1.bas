Attribute VB_Name = "Module1"
Public LangIXX As Long
Public IDDP As Integer
Public exitX As Boolean
Public data() As String
Public id() As Long
Public countX As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Public LB As ListBox
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Public Const LB_SETTABSTOPS = &H192
Public Const LB_ITEMFROMPOINT = &H1A9
Public Function LongToInt(ByVal value As Long) As Integer
CopyMemory LongToInt, ByVal VarPtr(value), 2
End Function
Public Function IntToLong(ByVal value As Integer) As Long
CopyMemory ByVal VarPtr(IntToLong), value, 2
End Function
Public Sub DeleteEntry(ByVal value As Long)
If countX = 1 Then
countX = 0
Erase data
Erase id
LB.Clear
Exit Sub
End If
If value = UBound(data) Then
ReDim Preserve data(UBound(data) - 1)
ReDim Preserve id(UBound(id) - 1)
LB.RemoveItem value
countX = countX - 1
Else
CopyMemory ByVal VarPtr(data(value)), ByVal VarPtr(data(value + 1)), (UBound(data) - value) * 4
CopyMemory ByVal VarPtr(data(UBound(data))), 0&, 4
CopyMemory ByVal VarPtr(id(value)), ByVal VarPtr(id(value + 1)), (UBound(id) - value) * 4
ReDim Preserve data(UBound(data) - 1)
ReDim Preserve id(UBound(id) - 1)
LB.RemoveItem value
countX = countX - 1
End If
End Sub

Public Sub PutData(ByVal IDX As Long, ByVal stringX As String)
Dim Mstrlen As Long
Mstrlen = 35
If countX = 0 Then
countX = countX + 1
ReDim data(0)
ReDim id(0)
data(0) = stringX
id(0) = IDX
If Len(stringX) < 35 Then Mstrlen = Len(stringX)
LB.AddItem IDX & vbTab & Left(stringX, Mstrlen)

Else
'Provjeri duplikate
Dim entr As Long
entr = EntrySearch(id, IDX)

If entr = -1 Then
ReDim Preserve data(UBound(data) + 1)
ReDim Preserve id(UBound(id) + 1)
data(UBound(data)) = stringX
id(UBound(id)) = IDX
If Len(stringX) < 35 Then Mstrlen = Len(stringX)
LB.AddItem IDX & vbTab & Left(stringX, Mstrlen)
QuickSortMe id, 0, UBound(id)
countX = countX + 1
Else
data(entr) = stringX
id(entr) = IDX
If Len(stringX) < 35 Then Mstrlen = Len(stringX)
LB.List(entr) = IDX & vbTab & Left(stringX, Mstrlen)
End If
End If
End Sub


Public Sub SwapLI(ByVal ind1 As Long, ByVal ind2 As Long)
Dim T_HOLD As String
T_HOLD = LB.List(ind1)
LB.List(ind1) = LB.List(ind2)
LB.List(ind2) = T_HOLD
End Sub

Public Sub SwapStrings(pbString1 As String, pbString2 As String)
Dim l_Hold As Long
CopyMemory l_Hold, ByVal VarPtr(pbString1), 4
CopyMemory ByVal VarPtr(pbString1), ByVal VarPtr(pbString2), 4
CopyMemory ByVal VarPtr(pbString2), l_Hold, 4
End Sub
Public Sub SwapID(ID1 As Long, ID2 As Long)
Dim v_Hold As Long
CopyMemory v_Hold, ID1, 4
CopyMemory ID1, ID2, 4
CopyMemory ID2, v_Hold, 4
End Sub


Public Sub QuickSortMe(varArray() As Long, Optional l_First As Long = -1, Optional l_Last As Long = -1)
Dim l_Low As Long
Dim l_Middle As Long
Dim l_High As Long
Dim v_Test As Long

If l_First < l_Last Then
l_Middle = (l_First + l_Last) / 2
v_Test = varArray(l_Middle)
l_Low = l_First
l_High = l_Last

Do

While varArray(l_Low) < v_Test
l_Low = l_Low + 1
Wend

While varArray(l_High) > v_Test
l_High = l_High - 1
Wend

If (l_Low <= l_High) Then
SwapStrings data(l_Low), data(l_High)
SwapID varArray(l_Low), varArray(l_High)
SwapLI l_Low, l_High

l_Low = l_Low + 1
l_High = l_High - 1
End If
Loop While (l_Low <= l_High)

If l_First < l_High Then
QuickSortMe varArray, l_First, l_High
End If

If l_Low < l_Last Then
QuickSortMe varArray, l_Low, l_Last
End If
End If
End Sub
Public Function EntrySearch(arr() As Long, ByVal vall As Long) As Long
For u = 0 To UBound(arr)
If vall = arr(u) Then EntrySearch = u: Exit Function
Next u
EntrySearch = -1
End Function
Public Sub PutRESheader()
'PRE-HEADER
Put #1, , CLng(0)
Put #1, , CLng(&H20)
Put #1, , CLng(&HFFFF&)
Put #1, , CLng(&HFFFF&)
Put #1, , CLng(0)
Put #1, , CLng(0)
Put #1, , CLng(0)
Put #1, , CLng(0)
'END OF PRE-HEADER
End Sub
Public Function NameType(ByVal TEMPNAME As String, ByVal TEMPTYPE As String, NameX2() As Byte, TypeX2() As Byte) As Long()
Dim tmpLNG() As Long
ReDim tmpLNG(1)
If Not IsNumeric(TEMPTYPE) Then
TypeX2 = StrConv(TEMPTYPE & Chr(CByte(0)), vbFromUnicode)
tmpLNG(0) = VarPtr(TypeX2(0))
Else
tmpLNG(0) = CLng(TEMPTYPE)
End If
If Not IsNumeric(TEMPNAME) Then
NameX2 = StrConv(TEMPNAME & Chr(CByte(0)), vbFromUnicode)
tmpLNG(1) = VarPtr(NameX2(0))
Else
tmpLNG(1) = CLng(TEMPNAME)
End If
NameType = tmpLNG
End Function
Public Function CalculateEntry(ByVal value As Long) As Integer
CalculateEntry = Int(value / 16) + 1
End Function
Public Function CalculateEntryID(ByVal value As Long) As Integer
CalculateEntryID = value Mod 16
End Function
Public Sub SaveRES(ByVal filename As String)
On Error GoTo dalje:
Open filename For Binary As #1
PutRESheader

Dim TypeX1 As Long
Dim NameX1 As Long
Dim NameX2() As Byte
Dim TypeX2() As Byte

Dim tmpX() As Byte
Dim tmpEnt As Integer



Dim ddt As New Collection
Dim seen As New Collection

Dim x As Long
Do While x <= UBound(data)

tmpEnt = CalculateEntry(id(x))
ddt.Add CalculateEntryID((id(x)))
seen.Add x

x = x + 1
Do While tmpEnt = CalculateEntry(id(x)) And x <= UBound(data)
ddt.Add CalculateEntryID((id(x)))
seen.Add x
x = x + 1
Loop

dalje:
If Err <> 0 Then On Error GoTo 0
tmpX = JoinEntries(ddt, seen)

TypeX1 = 6
NameX1 = IntToLong(tmpEnt)

PutHeadMem NameX1, NameX2, TypeX1, TypeX2, tmpX, False, 0
Set seen = Nothing
Set ddt = Nothing
Loop

Dim TmpTmp() As Byte
For u = 1 To RestDataX.Count
TmpTmp = RestDataX.Item(u)
Put #1, , TmpTmp
Erase TmpTmp
Next u

Close #1
End Sub
Public Function JoinEntries(ddt As Collection, seen As Collection) As Byte()
Dim tmpData() As String
ReDim tmpData(15)
Dim JJ As String
Dim EntIDX As Long
Dim cntX As Integer
Dim teh() As Byte
Dim lLen As Long
For u = 1 To ddt.Count
EntIDX = IntToLong(ddt.Item(u))
If EntIDX > cntX Then
FillEmpty cntX, EntIDX - 1, tmpData
cntX = EntIDX
GoTo puni

Else
puni:
lLen = LenB(data(CLng(seen.Item(u))))
ReDim teh(lLen + 2 - 1)
CopyMemory teh(0), LongToInt(lLen / 2), 2
CopyMemory teh(2), ByVal (StrPtr(data(seen.Item(u)))), lLen
tmpData(EntIDX) = StrConv(teh, vbUnicode)
cntX = cntX + 1
End If
Next u

If cntX < 15 Then FillEmpty cntX, 15, tmpData
JJ = Join(tmpData, "")
JoinEntries = StrConv(JJ, vbFromUnicode)
End Function
Public Sub FillEmpty(ByVal from1 As Integer, ByVal to1 As Integer, strX() As String)
For u = from1 To to1
strX(u) = Chr(CByte(0)) & Chr(CByte(0))
Next u
End Sub
Public Sub PutHeadMem(ByVal NameX1 As Long, NameX2() As Byte, ByVal TypeX1 As Long, TypeX2() As Byte, MEMCONT() As Byte, Optional OnlyLoadMem As Boolean, Optional LANGX As Integer)
Dim ResHedLen As Long 'Resource Header length
Dim nameQ As Boolean
Dim typeQ As Boolean
Dim Resst2 As Long
Dim Resst As Long
ResHedLen = 24
If (NameX1 < 0) Or (NameX1 > &HFFFF&) Then
ResHedLen = ResHedLen + (lstrlen(VarPtr(NameX2(0))) + 1) * 2
nameQ = True
Else
ResHedLen = ResHedLen + 4
End If
If (TypeX1 < 0) Or (TypeX1 > &HFFFF&) Then
ResHedLen = ResHedLen + (lstrlen(VarPtr(TypeX2(0))) + 1) * 2
typeQ = True
Else
ResHedLen = ResHedLen + 4
End If
Put #1, , UBound(MEMCONT) + 1 'Dužina resourcea
Resst = ResHedLen Mod 4
If Resst <> 0 Then
ResHedLen = ResHedLen + Resst
End If
Put #1, , ResHedLen
If typeQ Then
Dim UNI1 As String
ReDim Preserve TypeX2(UBound(TypeX2) - 1)
UNI1 = StrConv(TypeX2, vbUnicode)
UNI1 = StrConv(UNI1, vbUnicode)
Put #1, , UNI1
Put #1, , CInt(0)
Else
Put #1, , CInt(&HFFFF)
Put #1, , CInt(TypeX1)
End If
If nameQ Then
Dim UNI2 As String
ReDim Preserve NameX2(UBound(NameX2) - 1)
UNI2 = StrConv(NameX2, vbUnicode)
UNI2 = StrConv(UNI2, vbUnicode)
Put #1, , UNI2
Put #1, , CInt(0)
Else
Put #1, , CInt(&HFFFF)
Put #1, , LongToInt(NameX1) '*************
End If
If Resst <> 0 Then Put #1, , CInt(0)
Put #1, , CLng(0) 'Data Version
Put #1, , CInt(&H1030) 'Memory Flag
Put #1, , LANGX
Put #1, , CLng(0) 'Version
Put #1, , CLng(0) 'Characteristic
Put #1, , MEMCONT 'Put Memory Data
'Do While (LOF(1) Mod 4) <> 0
'Put #1, , CByte(0)
'Loop
If ((ResHedLen + UBound(MEMCONT) + 1) Mod 4) <> 0 Then
Resst2 = ResHedLen + UBound(MEMCONT) + 1
Do While (Resst2 Mod 4) <> 0
Put #1, , CByte(0)
Resst2 = Resst2 + 1
Loop
End If

End Sub


