Sub ColorMeL2R_T2B()
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
        MacScript ("delay(0.0001)")
        If nextCell.Column = 26 Then
            Set nextCell = nextCell.Offset(1, -24)
        Else
            Set nextCell = nextCell.Offset(0, 1)
        End If
    Next
End Sub


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
        MacScript ("delay(0.00001)")
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
        MacScript ("delay(0.0001)")
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
    side = 31
    With Range(Cells.Address)
        .Interior.ThemeColor = 1
        .ColumnWidth = 4.5
        .RowHeight = 29#
    End With
    Set cell = Spread(side, Range("b1"), 1, 0)
    Call RecSpiral(side - 1, cell)
End Sub

Sub RecSpiral(num, cell)
    Set cell = Spread(num, cell, 0, 1)
    Set cell = Spread(num, cell, -1, 0)
    If num >= 2 Then
        Set cell = Spread(num - 1, cell, 0, -1)
        Set cell = Spread(num - 1, cell, 1, 0)
        If num >= 3 Then
            Call RecSpiral((num - 2), cell)
        End If
    End If
End Sub

Function Spread(num, cell, x, y)
    For i = 1 To num
        Set cell = cell.Offset(x, y)
        ColorCell cell
    Next
    Set Spread = cell
End Function

Sub ColorCell(cell)
        With cell.Interior
            .ThemeColor = Random(4, 12)
            .TintAndShade = 0.2 + ((0.1) * Random(0, 5))
        End With
        MacScript ("delay(0.0001)")
End Sub

Function Random(min, max)
    Random = Application.WorksheetFunction.RandBetween(min, max)
End Function

