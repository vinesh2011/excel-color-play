Sub ColorMeL2R_T2B_Easy()
    Dim nextCell As Range
    Dim increment
    Dim nextColor
    increment = True
    With Range(Cells.Address)
        .Interior.ThemeColor = 1
        .ColumnWidth = 4.5
        .RowHeight = 29#
    End With
    Set nextCell = Application.Range("B2")
    For counter = 1 To 625
        nextColor = Random(4, 12)
        With nextCell.Interior
            .ThemeColor = nextColor
            .TintAndShade = 0.2 + ((0.1) * Random(0, 5))
        End With
        Sleep0 1
        If nextCell.Column = 26 Then
            Set nextCell = nextCell.Offset(1, -24)
        Else
            Set nextCell = nextCell.Offset(0, 1)
        End If
    Next
End Sub

Sub ColorMeL2R_T2B_Complex()
    Dim nextCell As Range
    With Range("a1:ab30")
        .Interior.ThemeColor = 1
        .ColumnWidth = 4.5
        .RowHeight = 29#
    End With
    For counter = 1 To 4999
        Do
            nextColor = Random(3, 12)
            Set nextCell = Application.Range("B2").Offset(Random(0, 25), Random(0, 25))
            If isMyColorSameAsMyNeighbours(nextColor, nextCell) Then
                Exit Do
            End If
            With nextCell.Interior
                .ThemeColor = nextColor
                .TintAndShade = 0.2 + ((0.1) * Random(0, 5))
            End With
            Sleep0()
        Loop While False
    Next
End Sub

Function isMyColorSameAsMyNeighbours(nextColor, nextCell)
'  Location of My Neighbors in related to me:
'    |-------------|---------------|--------------|
'    | -1, -1(A)   |   0,-1 (B)    |   1,-1 (C)   |
'    |-------------|---------------|--------------|
'    | -1, 0  (D)  |   0,0 (Me)    |   1,0  (E)   |
'    |-------------|---------------|--------------|
'    | -1, 1  (F)  |   0,1   (G)   |   1,0 (H)    |
'    |-------------|---------------|--------------|

   '  Only change color if not set 
   '         If nextCell.Interior.ThemeColor > 1 Then
   '             Exit Do
   '         End If
    isSameAsMyNeighbour = False
    If nextColor = nextCell.Offset(-1, -1).Interior.ThemeColor Then     ' A
        isSameAsMyNeighbour = True
    ElseIf nextColor = nextCell.Offset(0, -1).Interior.ThemeColor Then  ' B
        isSameAsMyNeighbour = True
    ElseIf nextColor = nextCell.Offset(1, -1).Interior.ThemeColor Then  ' C
        isSameAsMyNeighbour = True
    ElseIf nextColor = nextCell.Offset(-1, 0).Interior.ThemeColor Then  ' D
        isSameAsMyNeighbour = True
    ElseIf nextColor = nextCell.Offset(1, 0).Interior.ThemeColor Then   ' E
        isSameAsMyNeighbour = True
    ElseIf nextColor = nextCell.Offset(-1, 1).Interior.ThemeColor Then  ' F
        isSameAsMyNeighbour = True
    ElseIf nextColor = nextCell.Offset(0, 1).Interior.ThemeColor Then   ' G
        isSameAsMyNeighbour = True
    ElseIf nextColor = nextCell.Offset(1, 0).Interior.ThemeColor Then   ' H
        isSameAsMyNeighbour = True
    End If
End Function

Sub SparklingSquares1()
    Dim nextCell As Range
    Dim nextColor
    With Range(Cells.Address)
        .Interior.ThemeColor = 1
        .ColumnWidth = 4.5
        .RowHeight = 29#
    End With
    For counter = 1 To 5000
        nextColor = Random(3, 12)
        Set nextCell = Application.Range("B2").Offset(Random(0, 25), Random(0, 25))
        With nextCell.Interior
            .ThemeColor = nextColor
            .TintAndShade = 0.2 + ((0.1) * Random(0, 5))
        End With
        Sleep0 1
    Next
End Sub

Sub ColorMe_Snake_And_Ladders()
    Dim nextCell As Range
    Dim factor
    Dim nextColor
    factor = 1
    With Range(Cells.Address)
        .Interior.ThemeColor = 1
        .ColumnWidth = 4.5
        .RowHeight = 29#
    End With
    Set nextCell = Application.Range("B2")
    For counter = 1 To 625
        nextColor = Random(4, 12)
        With nextCell.Interior
            .ThemeColor = nextColor
            .TintAndShade = 0.2 + ((0.1) * Random(0, 5))
        End With
        Sleep0 1
        If nextCell.Column = 26 And factor = 1 Then
            Set nextCell = nextCell.Offset(1, 0)
            factor = -1
        ElseIf nextCell.Column = 2 And factor = -1 Then
            Set nextCell = nextCell.Offset(1, 0)
            factor = 1
        Else
            Set nextCell = nextCell.Offset(0, factor)
        End If
    Next
End Sub

Sub Spiral()
    Dim side
    side = 30
    With Range(Cells.Address)
        .Interior.ThemeColor = 1
        .ColumnWidth = 4.5
        .RowHeight = 29#
    End With
    Set cell = Spread(side, Range("b1"), 1, 0)
    Call RecSpiral(side, cell)
End Sub

Sub RecSpiral(num, cell)
    Set cell = Spread(num, cell, 0, 1)
    Set cell = Spread(num - 1, cell, -1, 0)
    If num >= 2 Then
        Set cell = Spread(num - 2, cell, 0, -1)
        Set cell = Spread(num - 3, cell, 1, 0)
        If num >= 3 Then
            Call RecSpiral((num - 4), cell)
        End If
    End If
End Sub

Function Spread(num, cell, x, y)
    Dim clr
    clr = Random(4, 12)
    For i = 1 To num - 1
        Set cell = cell.Offset(x, y)
        Call ColorCell(cell, clr)
    Next
    Set Spread = cell
End Function

Sub ColorCell(cell, clr)
    With cell.Interior
        .ThemeColor = clr
        .TintAndShade = 0.2 + ((0.1) * Random(0, 2))
    End With
    Sleep0 1
End Sub

Function Random(min, max)
    Random = Application.WorksheetFunction.RandBetween(min, max)
End Function

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Sleep0( millis)
    #If Mac Then
        MacScript ("delay(" & (millis/1000) & ")")
    #Else
        Sleep millis
    #End If
Sub
