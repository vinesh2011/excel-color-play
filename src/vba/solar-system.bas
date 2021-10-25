
Sub SolarSystem()
    Dim center1 As Range
    Dim fullRange As Range
    Set fullRange = Range("a1:zz676")
    fullRange.Interior.ThemeColor = 2
    fullRange.Interior.TintAndShade = 0.2
    fullRange.ColumnWidth = 0.11
    fullRange.RowHeight = 1.05
    Set center1 = Range("$ka$330")
    ' QtrCircle0 Range("$a$1"), 200, 6, 7, False
    ' QtrCircle0 Range("$a$1"), 70, 8, 7, True
    If True Then
        ' Circle0 Range("$ej$530"), 30, 1, 1
       '  Circle0 Range("$ga$150"), 53, 6, 3
        'Circle0 Range("$ha$270"), 48, 9, 1
        'Circle0 Range("$ii$590"), 45, 12, 1
        'Circle0 Range("$la$240"), 68, 10, 1
        'Circle0 Range("$na$570"), 64, 4, 1
        'Circle0 Range("$qp$390"), 52, 7, 1
        'Circle0 Range("$sa$130"), 76, 11, 1
        'Circle0 Range("$sz$500"), 83, 8, 1
        'Circle0 Range("$xa$370"), 77, 5, 1
        ' Circle0 Range("$xz$410"), 38, 3, 1
        Circle0 Range("$lz$410"), 180, 8, 1
    End If
End Sub

Sub Circle0(centerCell, rad, clr, boundary)
    Dim Val, Tint As Double
    Dim x, color
    x = 0
    For radius = rad To 0 Step -boundary
       For i = 0 To (radius)
            Val = Round(Sqr((radius * radius) - (i * i)), 0)
            Tint = 0.8 - (0.8 * radius / rad)
            color = clr
            If x Mod 2 = 0 Then
                color = clr - 2
                Tint = (Tint + 0.3) Mod 0.8
            End If
            x = x + 1
            With Range(centerCell.Offset(i, Val), centerCell.Offset(i, -Val)).Interior
                .ThemeColor = color
                .TintAndShade = Tint
            End With
            With Range(centerCell.Offset(-i, Val), centerCell.Offset(-i, -Val)).Interior
                .ThemeColor = color
                ' .color
                .TintAndShade = Tint
            End With
       Next
        ' MacScript ("delay(0.00001)")
    Next
'    Range(centerCell.Offset(rad, 0), centerCell.Offset(-1 * rad, 0)).Interior.ThemeColor = color + 1
'    Range(centerCell.Offset(0, rad), centerCell.Offset(0, -1 * rad)).Interior.ThemeColor = color + 1
End Sub

Sub QtrCircle0(centerCell, radius, color, boundary, reverseShading)
    Dim Val, Tint As Double
    For rad = radius To 0 Step -boundary
        Tint = (0.8 * rad / radius)
        If reverseShading = False Then Tint = 0.8 - Tint
        For i = 1 To (rad)
            Val = Round(Sqr((radius * rad) - (i * i)), 0)
            With Range(centerCell.Offset(i, Val), centerCell.Offset(i, 0)).Interior
                .ThemeColor = color
                .TintAndShade = Tint
            End With
        Next
        MacScript ("delay(0.00001)")
    Next
End Sub

Sub testShad()
    QtrCircle0 Range("$a$1"), 200, 10, 3, False
End Sub
Sub QtrCircle1(centerCell, radius, clr, boundary, reverseShading)
    Dim Val, Tint As Double
    Dim x, color
    x = 0
    For rad = radius To 0 Step -boundary
        color = clr
        Tint = (0.8 * rad / radius)
        If x Mod 2 = 0 Then
             color = clr - 1
            ' Tint = (0.8 * rad / radius)
        End If
        If reverseShading = False Then Tint = 0.8 - Tint
        x = x + 1
        For i = 1 To (rad)
            Val = Round(Sqr((radius * rad) - (i * i)), 0)
            With Range(centerCell.Offset(i, Val), centerCell.Offset(i, 0)).Interior
                .ThemeColor = color
                .TintAndShade = Tint
            End With
        Next
        MacScript ("delay(0.00001)")
    Next
End Sub
