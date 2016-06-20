Attribute VB_Name = "CreatingBoard"
Option Explicit

Public Const BOARD_HEIGHT As Long = 10
Public Const BOARD_WIDTH As Long = 10
Public Const STARTING_SQUARE As String = "begin"
Public Const ENDING_SQUARE As String = "finish"
Public Const BOARD_SQUARE_HEIGHT As Long = 50
Public Const BOARD_SQUARE_WIDTH As Long = 9


Public Sub MakeGameBoard()
    Dim boardSheet As Worksheet
    
    'I can't figure out how to handle this error properly
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Game Board").Delete
    Set boardSheet = ThisWorkbook.Sheets.Add
    boardSheet.Name = "Game Board"
    Application.DisplayAlerts = True
    
    Set boardSheet = ThisWorkbook.Sheets("Game Board")
    Dim topLeftSquare As Range
    
    Set topLeftSquare = UsersPick
    
    DrawBoard boardSheet, topLeftSquare
   
    NumberBoard boardSheet
    
        
End Sub

Private Sub DrawBoard(ByVal targetSheet As Worksheet, ByVal topLeftCell As Range)
    Dim boardRange As Range
    Set boardRange = targetSheet.Range(Cells(topLeftCell.Row, topLeftCell.Column), Cells(topLeftCell.Row + BOARD_WIDTH - 1, topLeftCell.Column + BOARD_HEIGHT - 1))
    boardRange.Name = "gameboard"
    With boardRange
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThick
        .Rows.RowHeight = BOARD_SQUARE_HEIGHT
        .Columns.ColumnWidth = BOARD_SQUARE_WIDTH
        .Cells(1, 1).Name = ENDING_SQUARE
        .Cells(BOARD_HEIGHT, 1).Name = STARTING_SQUARE
    End With
End Sub

Private Function UsersPick() As Range
Dim topLeftSquare As Range
Set topLeftSquare = Application.InputBox("Please select the top left cell for your board", "Place Board", Type:=8)
    If topLeftSquare.Row = 1 Then Set topLeftSquare = Cells(2, topLeftSquare.Column)
    If topLeftSquare.Column = 1 Then Set topLeftSquare = Cells(topLeftSquare.Row, 2)
    Set UsersPick = topLeftSquare
End Function

Private Sub NumberBoard(ByVal boardSheet As Worksheet)
    Dim squareNumber As Long
    squareNumber = 100
    Dim boardRow As Long
    Dim boardColumn As Long
    Dim gameArea As Range
    Dim fillRightToLeft As Boolean
    With boardSheet
        Set gameArea = Range("gameboard")
        For boardRow = 1 To gameArea.Rows.Count
          
          fillRightToLeft = boardRow Mod 2
          
          Select Case fillRightToLeft
          
            Case True
              For boardColumn = 1 To BOARD_WIDTH
                  gameArea.Cells(boardRow, boardColumn) = squareNumber
                  squareNumber = squareNumber - 1
              Next
              
            Case False
              For boardColumn = BOARD_WIDTH To 1 Step -1
                  gameArea.Cells(boardRow, boardColumn) = squareNumber
                  squareNumber = squareNumber - 1
              Next
            
          End Select
        Next
    End With
End Sub
