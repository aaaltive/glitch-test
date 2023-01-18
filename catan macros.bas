
' ***********************************************************************************************************
' Globals
' ***********************************************************************************************************



' ***********************************************************************************************************
'FUNCTIONS
' ***********************************************************************************************************

' An array of all the tile edges


Public Function GET_ALL_EDGES() As Variant

        GET_ALL_EDGES = Array("Straight Connector 116", _
        "Straight Connector 113", "Straight Connector 111", "Straight Connector 108", _
        "Straight Connector 98", "Straight Connector 95", "Straight Connector 99", _
        "Straight Connector 141", "Straight Connector 140", "Straight Connector 90", _
        "Straight Connector 94", "Straight Connector 92", "Straight Connector 136", _
        "Straight Connector 143", "Straight Connector 132", "Straight Connector 130", _
        "Straight Connector 129", "Straight Connector 125", "Straight Connector 124", _
        "Straight Connector 120", "Straight Connector 119", "Straight Connector 106", _
        "Straight Connector 142", "Straight Connector 105", "Straight Connector 59", _
        "Straight Connector 64", "Straight Connector 63", "Straight Connector 78", _
        "Straight Connector 144", "Straight Connector 107", "Straight Connector 117", _
        "Straight Connector 112", "Straight Connector 114", "Straight Connector 115", _
        "Straight Connector 109", "Straight Connector 110", "Straight Connector 96", _
        "Straight Connector 97", "Straight Connector 137", "Straight Connector 138", _
        "Straight Connector 139", "Straight Connector 58", "Straight Connector 77", _
        "Straight Connector 74", "Straight Connector 83", "Straight Connector 80", _
        "Straight Connector 88", "Straight Connector 85", "Straight Connector 93", _
        "Straight Connector 67", "Straight Connector 79", "Straight Connector 84", _
        "Straight Connector 89", "Straight Connector 60", "Straight Connector 75", _
        "Straight Connector 76", "Straight Connector 81", "Straight Connector 82", _
        "Straight Connector 86", "Straight Connector 87", "Straight Connector 91", _
        "Straight Connector 133", "Straight Connector 134", "Straight Connector 135", _
        "Straight Connector 121", "Straight Connector 118", "Straight Connector 126", _
        "Straight Connector 123", "Straight Connector 131", "Straight Connector 122", _
        "Straight Connector 127")
        
End Function

Private Function GET_ALL_INTERSECTIONS() As Variant

    GET_ALL_INTERSECTIONS = Array("Oval 3", "Oval 161", "Oval 168", "Oval 162", _
        "Oval 170", "Oval 163", "Oval 172", "Oval 171", "Oval 169", "Oval 167", _
        "Oval 174", "Oval 173", "Oval 206", "Oval 164", "Oval 165", "Oval 166", _
        "Oval 212", "Oval 214", "Oval 199", "Oval 198", "Oval 197", "Oval 208", _
        "Oval 207", "Oval 185", "Oval 179", "Oval 181", "Oval 183", "Oval 213", _
        "Oval 210", "Oval 182", "Oval 180", "Oval 178", "Oval 184", "Oval 205", _
        "Oval 204", "Oval 175", "Oval 176", "Oval 177", "Oval 209", "Oval 211", _
        "Oval 202", "Oval 201", "Oval 200", "Oval 203", "Oval 196", "Oval 190", _
        "Oval 192", "Oval 194", "Oval 193", "Oval 160", "Oval 191", "Oval 159", _
        "Oval 189", "Oval 195", "Oval 399")

End Function

' An array of all the game board tiles


Public Function GET_ALL_TILES() As Variant

        GET_ALL_TILES = Array("Tile 1", "Tile 2", "Tile 3", _
        "Tile 4", "Tile 5", "Tile 6", "Tile 7", "Tile 8", _
        "Tile 9", "Tile 10", "Tile 11", "Tile 12", "Tile 13", _
        "Tile 14", "Tile 15", "Tile 16", "Tile 17", "Tile 18", _
        "Tile 19")
        
End Function



Public Function GET_ALL_NUMBER_TILES() As Variant

        GET_ALL_NUMBER_TILES = Array("Oval 1", "Oval 2", "Oval 3", _
        "Oval 4", "Oval 5", "Oval 6", "Oval 7", "Oval 8", _
        "Oval 9", "Oval 10", "Oval 11", "Oval 12", "Oval 13", _
        "Oval 14", "Oval 15", "Oval 16", "Oval 17", "Oval 18", _
        "Oval 19")
        
End Function



Public Function GetRandomColor(colorCount() As Long) As Integer
    Dim totalCount As Long
    totalCount = 0
    For i = 1 To UBound(colorCount)
        totalCount = totalCount + colorCount(i)
    Next i

    Dim randomIndex As Integer
    randomIndex = Int((totalCount) * Rnd + 1)

    Dim countSoFar As Long
    countSoFar = 0
    For i = 1 To UBound(colorCount)
        countSoFar = countSoFar + colorCount(i)
        If randomIndex <= countSoFar Then
            GetRandomColor = i
            Exit Function
        End If
    Next i
End Function



Function Reset_Players_Stats()
'
' Reset_Players_Stats Macro
' Resets all players stats
'

'
    Range("M17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("N17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("O17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("P17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("Q17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("R17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("S17").Select
    ActiveCell.FormulaR1C1 = "-"
    Range("T17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("U17").Select
    ActiveCell.FormulaR1C1 = "-"
    Range("V17").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("W17").Select
    ActiveCell.FormulaR1C1 = "-"
    Range("M17").Select
    Selection.AutoFill Destination:=Range("M17:M20"), Type:=xlFillDefault
    Range("M17:M20").Select
    Range("N17").Select
    Selection.AutoFill Destination:=Range("N17:N20"), Type:=xlFillDefault
    Range("N17:N20").Select
    Range("O17").Select
    Selection.AutoFill Destination:=Range("O17:O20"), Type:=xlFillDefault
    Range("O17:O20").Select
    Range("P17").Select
    Selection.AutoFill Destination:=Range("P17:P20"), Type:=xlFillDefault
    Range("P17:P20").Select
    Range("Q17").Select
    Selection.AutoFill Destination:=Range("Q17:Q20"), Type:=xlFillDefault
    Range("Q17:Q20").Select
    Range("R17").Select
    Selection.AutoFill Destination:=Range("R17:R20"), Type:=xlFillDefault
    Range("R17:R20").Select
    Range("S17").Select
    Selection.AutoFill Destination:=Range("S17:S20"), Type:=xlFillDefault
    Range("S17:S20").Select
    Range("T17").Select
    Selection.AutoFill Destination:=Range("T17:T20"), Type:=xlFillDefault
    Range("T17:T20").Select
    Range("U17").Select
    Selection.AutoFill Destination:=Range("U17:U20"), Type:=xlFillDefault
    Range("U17:U20").Select
    Range("V17").Select
    Selection.AutoFill Destination:=Range("V17:V20"), Type:=xlFillDefault
    Range("V17:V20").Select
    Range("W17").Select
    Selection.AutoFill Destination:=Range("W17:W20"), Type:=xlFillDefault
    Range("W17:W20").Select
End Function



' ***********************************************************************************************************
' Subroutines
' ***********************************************************************************************************



Public Sub Set_Board()
    'Set_Board Macro
    'Sets up the board randomly and assign random colors
    '
    
    Sheets("Sheet2").Range("A3:E21").ClearContents

    
    Dim NUM_TILE_VALUES As Variant
    NUM_TILE_VALUES = Array("10", "2", "9", "10", "8", "5", "11", "6", "5", "8", "9", "12", "6", "4", "3", "4", "3", "11", "", "")
    
    Dim response As Integer
    response = MsgBox("Are you sure you want to reset the game?", vbYesNo + vbQuestion, "Reset Game")
    If response = vbYes Then
       ' proceed with the rest of the subroutine
        Dim colorRange As Range
        Set colorRange = ActiveSheet.Range("M3:Q3")
        Dim colorCount(1 To 6) As Long
        colorCount(1) = 4 ' M3 Terrain: FD
        colorCount(2) = 3 ' N3 Terrain: H
        colorCount(3) = 4 ' O3 Terrain: P
        colorCount(4) = 4 ' P3 Terrain: FT
        colorCount(5) = 3 ' Q3 Terrain: M
        colorCount(6) = 1 ' #CC9900 Terrain: D
    
        Dim shapes As ShapeRange
        Dim numberTiles As ShapeRange
        Set shapes = ActiveSheet.shapes.Range(GET_ALL_TILES())
        Set numberTiles = ActiveSheet.shapes.Range(GET_ALL_NUMBER_TILES())
        Dim i As Long
        Dim j As Long
        j = 0
        For i = 1 To shapes.Count
            Dim colorIndex As Integer
            colorIndex = GetRandomColor(colorCount)
            If colorIndex <= 5 Then
                shapes(i).Fill.ForeColor.RGB = colorRange.Item(colorIndex).Interior.Color
                ' update cell in sheet 2, add terrain
                numberTiles(i).TextFrame.Characters.Text = NUM_TILE_VALUES(j) ' fill the number tile with the corresponding value from NUM_TILE_VALUES
                ' update cell in sheet 2, add Value
                Sheets("Sheet2").Select
                    Select Case colorIndex
                        Case 1
                            ActiveSheet.Cells(i + 2, 1).Value = "FD"
                            ActiveSheet.Cells(i + 2, 2).Value = NUM_TILE_VALUES(j)

                        Case 2
                            ActiveSheet.Cells(i + 2, 1).Value = "H"
                            ActiveSheet.Cells(i + 2, 2).Value = NUM_TILE_VALUES(j)

                        Case 3
                            ActiveSheet.Cells(i + 2, 1).Value = "P"
                            ActiveSheet.Cells(i + 2, 2).Value = NUM_TILE_VALUES(j)

                        Case 4
                            ActiveSheet.Cells(i + 2, 1).Value = "FT"
                            ActiveSheet.Cells(i + 2, 2).Value = NUM_TILE_VALUES(j)

                        Case 5
                            ActiveSheet.Cells(i + 2, 1).Value = "M"
                            ActiveSheet.Cells(i + 2, 2).Value = NUM_TILE_VALUES(j)
                    End Select
                Sheets("Sheet1").Select
                j = j + 1
            Else
                shapes(i).Fill.ForeColor.RGB = RGB(212, 169, 126)
                numberTiles(i).TextFrame.Characters.Text = ""
                Sheets("Sheet2").Select
                ActiveSheet.Cells(i + 2, 1).Value = "D"
                Sheets("Sheet1").Select
            End If
            colorCount(colorIndex) = colorCount(colorIndex) - 1
        Next i

        Call clear_selected_edges_and_intersections
        Call Reset_Players_Stats
    Else
        ' do nothing
    End If

End Sub



Public Sub clear_selected_edges_and_intersections()
'
' clear_selected_edges_and_intersections Macro
' selects all edges and clears any selection
'

'
    ActiveSheet.shapes.Range(GET_ALL_EDGES()).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
    End With
    
    ActiveSheet.Range("A1").Select
    
    ActiveSheet.shapes.Range(GET_ALL_INTERSECTIONS()).Select
    With Selection.ShapeRange
        .Line.Weight = 0
    End With

End Sub



Public Sub Worksheet_BeforeDoubleClick(ByVal Target As Shape, Cancel As Boolean)
    If Target.Type = msoLine Then
        Call Select_edge(Target)
    End If
End Sub



Sub Select_edge()
    'Select_edge Macro
    'when an edge is clicked, it changes to white to show it is selected
    '
    Dim selectedShape As String, sh As Shape
    selectedShape = Application.Caller

    Call clear_selected_edges_and_intersections 'clear previous selection

    Set sh = ActiveSheet.shapes(selectedShape)
    With sh.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Sheets("Sheet2").Range("J2").Value = selectedShape
    ActiveSheet.Range("A1").Select

End Sub



Sub Select_intersection()


    'Select_intersection macro
    'when am intersection is clicked, it becomes highlighted by changing the outline of the circle at the intersection to white
    
    Dim selectedShape As String, sh As Shape
    selectedShape = Application.Caller
    
    Call clear_selected_edges_and_intersections 'clear previous selection
    
    Set sh = ActiveSheet.shapes(selectedShape)
    With sh.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Weight = 2
        .Transparency = 0
    End With
    Sheets("Sheet2").Range("J2").Value = selectedShape
    ActiveSheet.Range("A1").Select
    
End Sub


Sub Roll_Dice()
    'Roll_Dice Macro
    'Rolls the dice


    Dim redDiceGroups() As Variant
    redDiceGroups = Array("Rdice1", "Rdice2", "Rdice3", "Rdice4", "Rdice5", "Rdice6")

    Dim yellowDiceGroups() As Variant
    yellowDiceGroups = Array("Ydice1", "Ydice2", "Ydice3", "Ydice4", "Ydice5", "Ydice6")

    Dim redDiceIndex As Integer
    redDiceIndex = Int((6 - 1 + 1) * Rnd + 0)

    Dim yellowDiceIndex As Integer
    yellowDiceIndex = Int((6 - 1 + 1) * Rnd + 0)

    Sheets("Sheet2").Select
    ActiveSheet.shapes.Range(redDiceGroups(redDiceIndex)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveSheet.shapes.Range(Array("Rdice")).Select
    Selection.Delete
    Range("B6:B8").Select
    ActiveSheet.Paste
    Selection.ShapeRange.Name = "Rdice"
    Selection.Name = "Rdice"

    Sheets("Sheet2").Select
    ActiveSheet.shapes.Range(yellowDiceGroups(yellowDiceIndex)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveSheet.shapes.Range(Array("Ydice")).Select
    Selection.Delete
    Range("B9:B11").Select
    ActiveSheet.Paste
    Selection.ShapeRange.Name = "Ydice"
    Selection.Name = "Ydice"
    ActiveSheet.shapes.Range(Array("RollDiceButton")).Select
    Selection.ShapeRange.ZOrder msoBringToFron
    ActiveSheet.Range("A1").Select

End Sub



