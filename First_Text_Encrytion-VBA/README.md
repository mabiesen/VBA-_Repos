# FirstTextEncrytion-VBA
This repository contains my first attempt at text encryption in VBA.  The intention of the macro is to secure passwords.

```vba

Sub subsub()
Dim rng As Range
Dim prase As String

Dim won As String
Dim too As String
Dim thrie As String
Dim letr As String
Dim i As Long
Dim k As Long
Dim flippedchar As String





Set rng = Application.InputBox("Select a cell for conversion", "Obtain Range Object", Type:=8)
prase = rng.Value
i = 0
ans = MsgBox("Should this macro endrop?", vbYesNo, "WhatToDo")
If ans = vbYes Then
    GoTo Endropnow
Else
    GoTo unendropnow
End If



'Endrop
Endropnow:
rng.Value = resortdata(rng.Value)
flippedchar = rng.Value
Range("D2").Value = "*"
For i = 1 To (Len(flippedchar))
letr = ""
    If Mid(flippedchar, i, 1) = " " Then
        Range("D2").Value = Range("D2").Value & "_"
        GoTo thenexti
    ElseIf Mid(flippedchar, i, 1) = "_" Then
        Range("D2").Value = Range("D2").Value & "_"
        GoTo thenexti
    ElseIf Mid(flippedchar, i, 1) = "?" Then
        Range("D2").Value = Range("D2").Value & "?"
        GoTo thenexti
    ElseIf Mid(flippedchar, i, 1) = "." Then
        Range("D2").Value = Range("D2").Value & "."
        GoTo thenexti
    ElseIf Mid(flippedchar, i, 1) = "," Then
        Range("D2").Value = Range("D2").Value & ","
        GoTo thenexti
    Else
        letr = Mid(flippedchar, i, 1)
        Range("D2").Value = Range("D2").Value & Endroption1rev(letr)
        GoTo thenexti
    End If
thenexti:
Next i
GoTo theveryend

'Unendrop
unendropnow:
Range("D1").Value = "*"
For i = 1 To (Len(prase))
won = ""
too = ""
thrie = ""
letr = ""
    If Mid(prase, i, 1) = "_" Then
        Range("D1").Value = Range("D1").Value & " "
        GoTo thenexti2
    ElseIf Mid(prase, i, 1) = " " Then
        Range("D1").Value = Range("D1").Value & " "
        GoTo thenexti2
    ElseIf Mid(prase, i, 1) = "." Then
        Range("D1").Value = Range("D1").Value & "."
        GoTo thenexti2
    ElseIf Mid(prase, i, 1) = "?" Then
        Range("D1").Value = Range("D1").Value & "?"
        GoTo thenexti2
    ElseIf Mid(prase, i, 1) = "," Then
        Range("D1").Value = Range("D1").Value & ","
        GoTo thenexti2
    ElseIf Mid(prase, i, 1) = "-" Then
        won = Mid(prase, i, 1)
        too = Mid(prase, i + 1, 1)
        thrie = Mid(prase, i + 2, 1)
        letr = won + too + thrie
        Range("D1").Value = Range("D1").Value & Endroption1(letr)
        i = i + 2
        GoTo thenexti2
    ElseIf Mid(prase, i, 1) = "+" Then
        won = Mid(prase, i, 1)
        too = Mid(prase, i + 1, 1)
        thrie = Mid(prase, i + 2, 1)
        letr = won + too + thrie
        Range("D1").Value = Range("D1").Value & Endroption1(letr)
        i = i + 2
        GoTo thenexti2
    Else
    End If
    
thenexti2:

Next i


GoTo theveryend

theveryend:

If ans = vbNo Then
    Range("D1").Value = resortdata(Range("D1").Value)
    MsgBox ("You will now have 2 minutes to view the dendropped information")
    a = Timer()
    b = Timer()
    MsgBox (Time())
    Application.Wait (Now + TimeValue("0:02:00"))
    Range("D1").Value = "*"
    MsgBox ("Time's up! Run again if you need more time")
Else
End If


End Sub
Public Function Endroption1(v) As Variant

Select Case v
    Case "-21"
        Endroption1 = "a"
    Case "-11"
        Endroption1 = "b"
    Case "-01"
        Endroption1 = "c"
    Case "-90"
        Endroption1 = "d"
    Case "-80"
        Endroption1 = "e"
    Case "-70"
        Endroption1 = "f"
    Case "-60"
        Endroption1 = "g"
    Case "-50"
        Endroption1 = "h"
    Case "-40"
        Endroption1 = "i"
    Case "-30"
        Endroption1 = "j"
    Case "-20"
        Endroption1 = "k"
    Case "-10"
        Endroption1 = "l"
    Case "-00"
        Endroption1 = "m"
    Case "+10"
        Endroption1 = "n"
    Case "+20"
        Endroption1 = "o"
    Case "+30"
        Endroption1 = "p"
    Case "+40"
        Endroption1 = "q"
    Case "+50"
        Endroption1 = "r"
    Case "+60"
        Endroption1 = "s"
    Case "+70"
        Endroption1 = "t"
    Case "+80"
        Endroption1 = "u"
    Case "+90"
        Endroption1 = "v"
    Case "+01"
        Endroption1 = "w"
    Case "+11"
        Endroption1 = "x"
    Case "+21"
        Endroption1 = "y"
    Case "+31"
        Endroption1 = "z"
End Select

End Function
Public Function Endroption1rev(v) As Variant

Select Case v
    Case "a"
        Endroption1rev = "-21"
    Case "b"
        Endroption1rev = "-11"
    Case "c"
        Endroption1rev = "-01"
    Case "d"
        Endroption1rev = "-90"
    Case "e"
        Endroption1rev = "-80"
    Case "f"
        Endroption1rev = "-70"
    Case "g"
        Endroption1rev = "-60"
    Case "h"
        Endroption1rev = "-50"
    Case "i"
        Endroption1rev = "-40"
    Case "j"
        Endroption1rev = "-30"
    Case "k"
        Endroption1rev = "-20"
    Case "l"
        Endroption1rev = "-10"
    Case "m"
        Endroption1rev = "-00"
    Case "n"
        Endroption1rev = "+10"
    Case "o"
        Endroption1rev = "+20"
    Case "p"
        Endroption1rev = "+30"
    Case "q"
        Endroption1rev = "+40"
    Case "r"
        Endroption1rev = "+50"
    Case "s"
        Endroption1rev = "+60"
    Case "t"
        Endroption1rev = "+70"
    Case "u"
        Endroption1rev = "+80"
    Case "v"
        Endroption1rev = "+90"
    Case "w"
        Endroption1rev = "+01"
    Case "x"
        Endroption1rev = "+11"
    Case "y"
        Endroption1rev = "+21"
    Case "z"
        Endroption1rev = "+31"
End Select

End Function

Public Function resortdata(data) As String
    Dim First As Long, Last As Long, total As Long
    Dim y As Long, z As Long, x As Long
    Dim Temp As String
    Dim Characterset() As String
    Dim newstring As String
    Dim numchar As Long
    
    numchar = Len(data)
    
    ReDim Characterset(0 To numchar) As String
    
    
    For x = 1 To Len(data)
        Characterset(x) = Mid(data, x, 1)
    Next x
        
    total = Len(data)
    First = 0
    If Len(data) Mod 2 = 0 Then
        Last = Len(data) / 2
    Else
        Last = Round(Len(data) / 2, 0)
    End If
    
    For y = First To Last
        z = total - y
            If z > y Then
                Temp = Characterset(y)
                Characterset(y) = Characterset(z)
                Characterset(z) = Temp
            End If
    Next y
    
    For y = First To total
        newstring = newstring & Characterset(y)
    Next y
    
    resortdata = newstring
                
End Function

```
