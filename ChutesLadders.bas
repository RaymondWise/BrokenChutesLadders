Attribute VB_Name = "ChutesLadders"
Option Explicit

Public Const TOTAL_LADDER As Long = 100
Public Const TOTAL_CHUTE As Long = -150

Public Sub ChutesAndLadders()
    'testing only
    Range("A:O").Clear

    
    
    
    Dim allChutesLadders As Long
    Dim numberOfChutes As Long
    Dim numberOfLadders As Long
    Dim i As Long
    Dim chutes As Variant
    Dim ladders As Variant
    
    allChutesLadders = totalObjects
    numberOfChutes = Int(allChutesLadders * Rnd + 1)
    If numberOfChutes = allChutesLadders Then numberOfChutes = allChutesLadders - 1
    If numberOfChutes < 3 Then numberOfChutes = 3
    numberOfLadders = allChutesLadders - numberOfChutes
    
    If numberOfLadders < 3 Then
        numberOfLadders = 3
        numberOfChutes = allChutesLadders - numberOfLadders
    End If
        
    'For testing only; should be private subs
    Cells(1, 5) = "Total"
    Cells(1, 6) = "Chutes"
    Cells(1, 7) = "Ladders"
    Cells(2, 5) = allChutesLadders
    Cells(2, 6) = numberOfChutes
    Cells(2, 7) = numberOfLadders
    Cells(1, 1) = "Chutes"
    Cells(1, 2) = "Ladders"
    Range("F3").Formula = "=sum(A2:A15)"
    Range("G3").Formula = "=sum(B2:B15)"
    Range("e3").Formula = "=sum(F3:G3)"
    Range("e4").Formula = "=sum(F4:G4)"
    
    
    
    chutes = ChuteLadderLengths(numberOfChutes, TOTAL_CHUTE, True)
        
        For i = 1 To UBound(chutes)
            Cells(i + 1, 1) = chutes(i)
        Next i

    
    ladders = ChuteLadderLengths(numberOfLadders, TOTAL_LADDER, False)

        For i = 1 To UBound(ladders)
            Cells(i + 1, 2) = ladders(i)
        Next i
    
    Dim ladderlocations As Variant
    Dim chutelocations As Variant
    ladders = GetLocations(ladders, False)
    chutes = GetLocations(chutes, True)
    
    'testing
    Cells(1, 9) = "chute begin"
    Cells(1, 10) = "chute end"
    Cells(1, 11) = "chute delta"
    Cells(1, 13) = "ladder begin"
    Cells(1, 14) = "ladder end"
    Cells(1, 15) = "ladder delta"
    
    
    FixRules chutes, ladders
    
    For i = 1 To UBound(ladders)
        Cells(i + 1, 13) = ladders(i, 1)
        Cells(i + 1, 14) = ladders(i, 2)
        Cells(i + 1, 15) = ladders(i, 2) - ladders(i, 1)
    Next
    
    For i = 1 To UBound(chutes)
        Cells(i + 1, 9) = chutes(i, 1)
        Cells(i + 1, 10) = chutes(i, 2)
        Cells(i + 1, 11) = chutes(i, 2) - chutes(i, 1)
    Next
    Cells(4, 6).Formula = "=Sum(K2:K20)"
    Cells(4, 7).Formula = "=Sum(O2:O20)"
    highlight
    Columns("A:O").AutoFit
End Sub

Private Function totalObjects() As Long
    Dim totalCount As Long
    totalCount = Int((17 - 9 + 1) * Rnd + 9)
    totalObjects = totalCount
End Function



Private Function ChuteLadderLengths(ByVal countChutesLadders As Long, ByVal totalChutesLadders As Long, ByVal isChute As Boolean) As Variant
    Dim index As Long
    Dim sumOfChutesLadders As Double
    Dim differenceFromTarget As Long
    Dim makeChutesNegative As Long
    makeChutesNegative = 1
    If isChute Then makeChutesNegative = -1
    
    Dim myChutesLadders() As Variant
    ReDim myChutesLadders(1 To countChutesLadders)
TryAgain:
    
    For index = 1 To countChutesLadders
        myChutesLadders(index) = Rnd()
    Next index
        
    sumOfChutesLadders = Application.WorksheetFunction.Sum(myChutesLadders)
    
    For index = 1 To countChutesLadders
        myChutesLadders(index) = Int(myChutesLadders(index) / sumOfChutesLadders * totalChutesLadders)
        
        If myChutesLadders(index) = 0 Then myChutesLadders(index) = makeChutesNegative * 2
        
    Next index
    
    sumOfChutesLadders = Application.WorksheetFunction.Sum(myChutesLadders)
    
    
    differenceFromTarget = totalChutesLadders - sumOfChutesLadders
    If differenceFromTarget <> 0 Then
        
        myChutesLadders(countChutesLadders) = myChutesLadders(countChutesLadders) + differenceFromTarget

    End If
    

    
    For index = 1 To countChutesLadders - 1
        If Abs(myChutesLadders(index)) >= 98 Then
            myChutesLadders(index) = myChutesLadders(index) - (makeChutesNegative * 50)
            myChutesLadders(countChutesLadders) = myChutesLadders(countChutesLadders) + (makeChutesNegative * 50)
        End If
    Next
    
    'why can chutes end with a positive number?
    If isChute And myChutesLadders(countChutesLadders) >= 0 Then
        myChutesLadders(countChutesLadders - 1) = myChutesLadders(countChutesLadders - 1) + myChutesLadders(countChutesLadders)
        countChutesLadders = countChutesLadders - 1
        ReDim Preserve myChutesLadders(1 To countChutesLadders)
    End If
    
    'Something can go wrong here
    If Abs(myChutesLadders(countChutesLadders)) >= 98 Then
        ReDim Preserve myChutesLadders(1 To countChutesLadders + 1)
        myChutesLadders(countChutesLadders + 1) = Application.WorksheetFunction.RoundDown(myChutesLadders(countChutesLadders) / 2, 0)
        myChutesLadders(countChutesLadders) = Application.WorksheetFunction.RoundUp(myChutesLadders(countChutesLadders) / 2, 0)

    End If
    
'HOW DO I HAVE POSITIVE CHUTES HERE, taking a shortcut
If isChute And myChutesLadders(countChutesLadders) >= 0 Then GoTo TryAgain

    ChuteLadderLengths = myChutesLadders()
            
End Function

Private Function GetLocations(ByVal items As Variant, ByVal isChute As Boolean) As Variant
    Dim totalItems As Long
    totalItems = UBound(items)
    
    Dim itemPositions As Variant
    ReDim itemPositions(1 To totalItems, 1 To 2)
    Dim index As Long
    Dim startingPoint As Long
    Dim objectSize As Long

    
    For index = 1 To totalItems

        startingPoint = 500
        objectSize = items(index)
        If isChute Then objectSize = objectSize * -1
        startingPoint = Int((99 - objectSize - 2 + 1) * Rnd + 2)
        If Not isChute Then itemPositions(index, 1) = startingPoint
        If Not isChute Then itemPositions(index, 2) = startingPoint + objectSize
        If isChute Then itemPositions(index, 2) = startingPoint
        If isChute Then itemPositions(index, 1) = startingPoint + objectSize
    Next
    
    
    GetLocations = itemPositions
    
    
End Function


Sub FixRules(ByVal chutesArray As Variant, ByVal laddersArray As Variant)
    Const myMin As Long = 2
    Const myMax As Long = 99
    Dim itemArray As Variant
    ReDim itemArray(1 To UBound(chutesArray) + UBound(laddersArray))
    
    Dim index As Long
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    
    For index = 1 To UBound(chutesArray)
FirstChance:
        If dict.Exists(CStr(chutesArray(index, 2))) Then
            'MsgBox "double chute"
            chutesArray(index, 1) = Int(chutesArray(index, 1)) + 1
            chutesArray(index, 2) = Int(chutesArray(index, 2)) + 1
            GoTo FirstChance
            
        End If
    dict(CStr(chutesArray(index, 2))) = 1
    Next

    For index = 1 To UBound(laddersArray)
SecondChance:
        If dict.Exists(CStr(laddersArray(index, 1))) Then
            'MsgBox "double ladder or double start"
            laddersArray(index, 1) = Int(laddersArray(index, 1)) + 1
            laddersArray(index, 2) = Int(laddersArray(index, 2)) + 1
            GoTo SecondChance
        End If
    dict(CStr(laddersArray(index, 1))) = 1
    Next

Cells(2, 9) = Application.Transpose(chutesArray)
Cells(2, 13) = Application.Transpose(laddersArray)

End Sub

Private Sub highlight()
Dim chuteLastRow As Long
Dim ladderLastRow As Long
chuteLastRow = Cells(Rows.Count, "I").End(xlUp).Row
ladderLastRow = Cells(Rows.Count, "M").End(xlUp).Row
Dim i As Long
Dim j As Long
Dim firstLookupValue As Long
Dim secondLookupValue As Long

For i = 2 To chuteLastRow
    firstLookupValue = Cells(i, "I")
    secondLookupValue = Cells(i, "J")
    For j = 2 To ladderLastRow
        If firstLookupValue = Cells(j, "N") Then
            Cells(i, "I").Interior.ColorIndex = 3
            Cells(j, "N").Interior.ColorIndex = 3
        End If
        If secondLookupValue = Cells(j, "M") Then
            Cells(i, "J").Interior.ColorIndex = 4
            Cells(j, "M").Interior.ColorIndex = 4
        End If
    Next
Next


End Sub
