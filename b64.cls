VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "b64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Function d2b(ip As Integer) As String
Dim q As Integer, r As Integer
r = 0
d2b = ""
If ip = 0 Then
d2b = String(8, "0")
Exit Function
End If
Do While ip <> 0
d2b = CStr(ip Mod 2) + d2b
ip = ip \ 2
Loop
d2b = String(8 - Len(d2b), "0") + d2b
End Function
Private Function b2d(ip As String) As Integer
j = 0
b2d = 0
For i = Len(ip) To 1 Step -1
If CInt(Mid(ip, i, 1)) = 1 Then
b2d = b2d + 2 ^ j
End If
j = j + 1
Next
End Function
Private Function enc(ip As Integer) As String
Select Case ip
Case 0: enc = "A"
Case 1: enc = "B"
Case 2: enc = "C"
Case 3: enc = "D"
Case 4: enc = "E"
Case 5: enc = "F"
Case 6: enc = "G"
Case 7: enc = "H"
Case 8: enc = "I"
Case 9: enc = "J"
Case 10: enc = "K"
Case 11: enc = "L"
Case 12: enc = "M"
Case 13: enc = "N"
Case 14: enc = "O"
Case 15: enc = "P"
Case 16: enc = "Q"
Case 17: enc = "R"
Case 18: enc = "S"
Case 19: enc = "T"
Case 20: enc = "U"
Case 21: enc = "V"
Case 22: enc = "W"
Case 23: enc = "X"
Case 24: enc = "Y"
Case 25: enc = "Z"
Case 26: enc = "a"
Case 27: enc = "b"
Case 28: enc = "c"
Case 29: enc = "d"
Case 30: enc = "e"
Case 31: enc = "f"
Case 32: enc = "g"
Case 33: enc = "h"
Case 34: enc = "i"
Case 35: enc = "j"
Case 36: enc = "k"
Case 37: enc = "l"
Case 38: enc = "m"
Case 39: enc = "n"
Case 40: enc = "o"
Case 41: enc = "p"
Case 42: enc = "q"
Case 43: enc = "r"
Case 44: enc = "s"
Case 45: enc = "t"
Case 46: enc = "u"
Case 47: enc = "v"
Case 48: enc = "w"
Case 49: enc = "x"
Case 50: enc = "y"
Case 51: enc = "z"
Case 52: enc = "0"
Case 53: enc = "1"
Case 54: enc = "2"
Case 55: enc = "3"
Case 56: enc = "4"
Case 57: enc = "5"
Case 58: enc = "6"
Case 59: enc = "7"
Case 60: enc = "8"
Case 61: enc = "9"
Case 62: enc = "+"
Case 63: enc = "/"
End Select
End Function
Public Function encode(es As String) As String
res = ""
j = 1
For i = 1 To Len(es) Step 3
g = Mid(es, i, 3)
na = 3 - Len(g)
g = g + String(na, Chr(0))
g = d2b(Asc(Mid(g, 1, 1))) + d2b(Asc(Mid(g, 2, 1))) + d2b(Asc(Mid(g, 3, 1)))
g = enc(b2d(Mid(g, 1, 6))) + enc(b2d(Mid(g, 7, 6))) + enc(b2d(Mid(g, 13, 6))) + enc(b2d(Mid(g, 19, 6)))
g = Left(g, 4 - na) + Left("==", na)
encode = encode + g
j = j + 1
If j = 20 Then
encode = encode + vbCrLf
j = 1
End If
Next
End Function
Private Function dec(ip As String) As Integer
Select Case ip
Case "A": dec = 0
Case "B": dec = 1
Case "C": dec = 2
Case "D": dec = 3
Case "E": dec = 4
Case "F": dec = 5
Case "G": dec = 6
Case "H": dec = 7
Case "I": dec = 8
Case "J": dec = 9
Case "K": dec = 10
Case "L": dec = 11
Case "M": dec = 12
Case "N": dec = 13
Case "O": dec = 14
Case "P": dec = 15
Case "Q": dec = 16
Case "R": dec = 17
Case "S": dec = 18
Case "T": dec = 19
Case "U": dec = 20
Case "V": dec = 21
Case "W": dec = 22
Case "X": dec = 23
Case "Y": dec = 24
Case "Z": dec = 25
Case "a": dec = 26
Case "b": dec = 27
Case "c": dec = 28
Case "d": dec = 29
Case "e": dec = 30
Case "f": dec = 31
Case "g": dec = 32
Case "h": dec = 33
Case "i": dec = 34
Case "j": dec = 35
Case "k": dec = 36
Case "l": dec = 37
Case "m": dec = 38
Case "n": dec = 39
Case "o": dec = 40
Case "p": dec = 41
Case "q": dec = 42
Case "r": dec = 43
Case "s": dec = 44
Case "t": dec = 45
Case "u": dec = 46
Case "v": dec = 47
Case "w": dec = 48
Case "x": dec = 49
Case "y": dec = 50
Case "z": dec = 51
Case "0": dec = 52
Case "1": dec = 53
Case "2": dec = 54
Case "3": dec = 55
Case "4": dec = 56
Case "5": dec = 57
Case "6": dec = 58
Case "7": dec = 59
Case "8": dec = 60
Case "9": dec = 61
Case "+": dec = 62
Case "/": dec = 63
Case "=": dec = 0
End Select
End Function

Public Function decode(ds As String) As String
On Error Resume Next
decode = ""
ds = Replace(ds, vbCr, "")
ds = Replace(ds, vbLf, "")
ds = Replace(ds, vbCrLf, "")
For i = 1 To Len(ds) Step 4
g = Mid(ds, i, 4)
g = Right(d2b(dec(Mid(g, 1, 1))), 6) + Right(d2b(dec(Mid(g, 2, 1))), 6) + Right(d2b(dec(Mid(g, 3, 1))), 6) + Right(d2b(dec(Mid(g, 4, 1))), 6)
g = Chr(b2d(Mid(g, 1, 8))) + Chr(b2d(Mid(g, 9, 8))) + Chr(b2d(Mid(g, 17, 8)))
decode = decode + g
Next
End Function

Private Function ba2s(ba() As Byte) As String
ba2s = ""
For i = 0 To UBound(ba)
ba2s = ba2s & Chr(ba(i))
Next
End Function
Private Function s2ba(s As String) As Byte()
Dim ba() As Byte
ReDim ba(Len(s) - 1) As Byte
For i = 0 To Len(s) - 1
ba(i) = Asc(Mid(s, i + 1))
Next
s2ba = ba
End Function

Private Function rf(ByVal fn As String) As String
On Error Resume Next
Dim fc As String
Dim fs As Long
Dim f As Integer
f = FreeFile
fs = FileLen(fn)
fc = String(fs, " ")
Open fn For Binary As #f
Get #f, , fc
Close #f
rf = fc
End Function

Public Sub encodefile(ip As String, op As String)
Dim fc() As Byte
fc = s2ba(encode(rf(ip)))
f = FreeFile
Open op For Binary Access Write As #f
Put #f, , fc
Close #f
End Sub
Public Sub decodefile(ip As String, op As String)
Dim fc() As Byte
fc = s2ba(decode(rf(ip)))
f = FreeFile
Open op For Binary Access Write As #f
Put #f, , fc
Close #f
End Sub

