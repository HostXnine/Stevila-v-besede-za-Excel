Attribute VB_Name = "Module1"
Option Explicit

Function SpellNumber(ByVal numIn)
    Dim LSide, RSide, Temp, DecPlace, Count, oNum
    oNum = numIn
    ReDim Place(9) As String
    Place(2) = "tisoè"
    Place(3) = "milijon"
    Place(4) = "milijard"
    Place(5) = "bilijon"
    numIn = Trim(Str(numIn)) 'String representation of amount
    ' Edit 2.(0)/Internationalisation
    ' Don't change point sign here as the above assignment preserves the point!
    DecPlace = InStr(numIn, ".") 'Pos of dec place 0 if none
    If DecPlace > 0 Then 'Convert Right & set numIn
        RSide = GetTens(Left(Mid(numIn, DecPlace + 1) & "00", 2))
        numIn = Trim(Left(numIn, DecPlace - 1))
    End If
    RSide = numIn 'converts numIn into words
    Count = 1
    Do While numIn <> ""
        Temp = GetHundreds(Right(numIn, 3))
        Select Case Temp = GetHundreds(Right(numIn, 3))
            Case Temp = "ena" And Count = 2 'da izpiše tisoè in ne ena tisoè
                LSide = Place(Count) & " " & LSide
            Case Temp = "ena" And Count = 3 'da izpiše milijon in ne ena milijon
                LSide = Place(Count) & " " & LSide
            Case Temp = "dva" And Count = 3 'sklanjatve za milijon
                LSide = Temp & " " & Place(Count) & "a " & LSide
            Case Temp = "tri" Or Temp = "štiri" And Count = 3
                LSide = Temp & " " & Place(Count) & "e " & LSide
            Case Temp <> "" And Count = 3
                LSide = Temp & " " & Place(Count) & "ov " & LSide
            Case Temp <> ""
                LSide = Temp & " " & Place(Count) & " " & LSide 'za vsa ostala števila ki sledijo sklanjatvam
        End Select
        If Len(numIn) > 3 Then
            numIn = Left(numIn, Len(numIn) - 3)
        Else
            numIn = ""
        End If
        Count = Count + 1
    Loop
    
    LSide = Trim(LSide)
    SpellNumber = LSide
    If InStr(oNum, Application.DecimalSeparator) > 0 Then    ' << Edit 2.(1) adding words for decimals
        SpellNumber = SpellNumber & " " & fractionWords(oNum) & "/100 "
    End If
    
End Function

Function GetHundreds(ByVal numIn) 'Converts a number from 100-999 into text
    Dim w As String
    If Val(numIn) = 0 Then Exit Function
    numIn = Right("000" & numIn, 3)
    If Mid(numIn, 1, 1) <> "0" Then 'Convert hundreds place with a few exceptions for sloveninan language
        Select Case GetDigit(Mid(numIn, 1, 1))
        Case "dva": w = "dvesto "
        Case "ena": w = "sto "
        Case Else: w = GetDigit(Mid(numIn, 1, 1)) & "sto "
        End Select
    End If
    If Mid(numIn, 2, 1) <> "0" Then 'Convert tens and ones place
        w = w & GetTens(Mid(numIn, 2))
    Else
        w = w & GetDigit(Mid(numIn, 3))
    End If
    w = Trim(w)
    GetHundreds = w
End Function

Function GetTens(TensText)  'Converts a number from 10 to 99 into text
    Dim w As String
    w = ""           'Null out the temporary function value
    If Val(Left(TensText, 1)) = 1 Then   'If value between 10-19
        Select Case Val(TensText)
            Case 10: w = "deset"
            Case 11: w = "enajst"
            Case 12: w = "dvanajst"
            Case 13: w = "trinajst"
            Case 14: w = "štirinajst"
            Case 15: w = "petnajst"
            Case 16: w = "šestnajst"
            Case 17: w = "sedemnajst"
            Case 18: w = "osemnajst"
            Case 19: w = "devetnajst"
            Case Else
        End Select
    Else      'If value between 20-99..
        Select Case Val(Left(TensText, 1))
            Case 2: w = "dvajset"
            Case 3: w = "trideset"
            Case 4: w = "štirideset"
            Case 5: w = "petdeset"
            Case 6: w = "šestdeset"
            Case 7: w = "sedemdeset"
            Case 8: w = "osemdeset"
            Case 9: w = "devetdeset"
            Case Else
        End Select
    End If
    If Val(Left(TensText, 1)) = 1 Or Val(Right(TensText, 1)) = 0 Then 'èe gre za število ki ni 11-19 ali da niso desetice vrini besedo "in"
        w = w
    Else
        w = GetDigit(Right(TensText, 1)) & "in" & w
    End If
    GetTens = w
End Function

Function GetDigit(Digit) 'Converts a number from 1 to 9 into text
    Select Case Val(Digit)
        Case 1: GetDigit = "ena"
        Case 2: GetDigit = "dva"
        Case 3: GetDigit = "tri"
        Case 4: GetDigit = "štiri"
        Case 5: GetDigit = "pet"
        Case 6: GetDigit = "šest"
        Case 7: GetDigit = "sedem"
        Case 8: GetDigit = "osem"
        Case 9: GetDigit = "devet"
        Case Else: GetDigit = ""
    End Select
End Function

Function fractionWords(n) As String
    Dim fraction As String, x As Long
    fraction = Split(n, Application.DecimalSeparator)(1)   ' << Edit 2.(2)
    For x = 1 To 2
        If fractionWords <> "" Then fractionWords = fractionWords & ""
        fractionWords = fractionWords & Mid(fraction, x, 1)
    Next x
    If Len(fractionWords) = 1 Then fractionWords = fractionWords & "0"
End Function

