Attribute VB_Name = "Module2"
Public F2CAP As String

Public Type RTSTRx
id As Long
data As String
End Type

Dim CHECKVAL As Long
Public RestDataX As New Collection
Public StringDataX As New Collection
Public StringIDX As New Collection
Dim CHECKINT As Integer
Public ResData() As Byte


Public Sub InsertDataX()
Dim Mstrlen As Long
Dim stringX As String
If countX = 0 Then Exit Sub
ReDim data(StringDataX.Count - 1)
ReDim id(StringIDX.Count - 1)
LB.Clear

For u = 1 To StringDataX.Count
data(u - 1) = StringDataX.Item(u)
id(u - 1) = StringIDX.Item(u)
stringX = StringDataX.Item(u)
Mstrlen = 35
If Len(stringX) < 35 Then Mstrlen = Len(stringX)
LB.AddItem (StringIDX.Item(u)) & vbTab & Left(stringX, Mstrlen)
Next u

QuickSortMe id, 0, UBound(id)

End Sub



Public Function EnumRESFile() As Boolean 'False ako je krivi file!
countX = 0
Set RestDataX = Nothing
Set StringDataX = Nothing
Set StringIDX = Nothing

Dim tempCNTX As Long
Dim TmpBXdata() As Byte

Dim lLen As Long
Dim countXX As Long
Dim VALSTRUC As Variant
Dim ext As Long
Dim SRClen As Long
Dim STRClen As Long
Dim startSTRC As Long
Dim tmpLong As Long

ext = UBound(ResData) + 1
VALSTRUC = Array("0", CStr(CLng(&H20)), CStr(CLng(&HFFFF&)), CStr(CLng(&HFFFF&)), "0", "0", "0", "0")
For u = 0 To 7
CopyMemory CHECKVAL, ResData(countXX), 4
If CHECKVAL <> CLng(VALSTRUC(u)) Then Exit Function
countXX = countXX + 4
Next u
Erase VALSTRUC

Do While countXX < ext
'Uzmi dužinu resourcea

startSTRC = countXX

CopyMemory SRClen, ResData(countXX), 4 '***
countXX = countXX + 4
CopyMemory STRClen, ResData(countXX), 4 '***
countXX = countXX + 4

'Uzmi TYPE
CopyMemory CHECKINT, ResData(countXX), 2

If CHECKINT = CInt(&HFFFF) Then
CopyMemory CHECKINT, ResData(countXX + 2), 2

If CHECKINT = 6 Then
GoTo dalje1
Else
GoTo copyelse
End If

Else
copyelse:
Dim kcbt As Byte
kcbt = 0
tmpLong = startSTRC + SRClen + STRClen
Do While (tmpLong Mod 4) <> 0
tmpLong = tmpLong + 1
kcbt = kcbt + 1
Loop
ReDim TmpBXdata(SRClen + STRClen - 1 + kcbt)
CopyMemory TmpBXdata(0), ResData(startSTRC), SRClen + STRClen
RestDataX.Add TmpBXdata

countXX = tmpLong: GoTo eend
End If

dalje1:
countXX = countXX + 4 + 2 'Preskoci &HFFFF-jer to sigurno dolazi

Dim entriesYY() As RTSTRx
Dim EntrYY1 As Integer
CopyMemory CHECKINT, ResData(countXX), 2
countXX = countXX + 2
'Dobili smo Entry
tmpLong = startSTRC + SRClen + STRClen
Do While (tmpLong Mod 4) <> 0
tmpLong = tmpLong + 1
Loop
ReDim TmpBXdata(SRClen + 1)
CopyMemory TmpBXdata(0), ResData(startSTRC + STRClen), tmpLong - startSTRC - STRClen

entriesYY = SetStringNameEx(TmpBXdata, CHECKINT)

For u = 0 To UBound(entriesYY)
StringDataX.Add entriesYY(u).data
StringIDX.Add entriesYY(u).id
countX = countX + 1
Next u

If (countXX Mod 4) <> 0 Then countXX = countXX + 2
countXX = countXX + 6 'Preskoci DataVersion i MemoryFlag
'Uzmi LangID
CopyMemory LangIXX, ResData(countXX), 2

countXX = tmpLong

eend:
Loop
Erase ResData
Erase TmpBXdata
EnumRESFile = True
End Function
Public Function SetStringNameEx(data() As Byte, ByVal entry As Long) As RTSTRx()
On Error GoTo eRe
Dim tmpBFR() As RTSTRx
ReDim tmpBFR(15)
Dim CountY As Integer
Dim countXX As Long
Dim CHECKLNG As Integer
Dim CHECKLNG1 As Long
entry = (entry - 1) * 16
For u = 0 To 15
CopyMemory CHECKLNG, data(countXX), 2
CHECKLNG1 = IntToLong(CHECKLNG)
If CHECKLNG1 = 0 Then countXX = countXX + 2: GoTo dalje
'If CHECKLNG1 > ResTotLen Then GoTo eRe
tmpBFR(CountY).data = Space(CHECKLNG1)
'Kopiraj UNICODE sadržaj
CopyMemory ByVal StrPtr(tmpBFR(CountY).data), data(countXX + 2), CHECKLNG1 * 2
countXX = countXX + 2 + CHECKLNG1 * 2
tmpBFR(CountY).id = entry
CountY = CountY + 1
dalje:
entry = entry + 1
Next u

ReDim Preserve tmpBFR(CountY - 1)
SetStringNameEx = tmpBFR
Erase tmpBFR
'Ova metoda je puno bolja jer sa LoadString ne znamo koje je velièine nadolazeci string,pa
'ne moramo puniti string proizvoljnom velicinom!
Exit Function
eRe:
On Error GoTo 0
ERRORX = True
End Function

