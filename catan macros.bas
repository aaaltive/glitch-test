
' ***********************************************************************************************************
' Globals
' ***********************************************************************************************************



' ***********************************************************************************************************
'FUNCTIONS
' ***********************************************************************************************************

' An array of all the tile edges


Public Function GET_ALL_EDGES() As Variant

        GET_ALL_EDGES = Array("Edge1", "Edge2", "Edge3", "Edge4", "Edge5", "Edge6", _
        "Edge7", "Edge8", "Edge9", "Edge10", "Edge11", "Edge12", "Edge13", "Edge14", _
        "Edge15", "Edge16", "Edge17", "Edge18", "Edge19", "Edge20", "Edge21", "Edge22", _
        "Edge23", "Edge24", "Edge25", "Edge26", "Edge27", "Edge28", "Edge29", "Edge30", _
        "Edge31", "Edge32", "Edge33", "Edge34", "Edge35", "Edge36", "Edge37", "Edge38", _
        "Edge39", "Edge40", "Edge41", "Edge42", "Edge43", "Edge44", "Edge45", "Edge46", _
        "Edge47", "Edge48", "Edge49", "Edge50", "Edge51", "Edge52", "Edge53", "Edge54", _
        "Edge55", "Edge56", "Edge57", "Edge58", "Edge59", "Edge60", "Edge61", "Edge62", _
        "Edge63", "Edge64", "Edge65", "Edge66", "Edge67", "Edge68", "Edge69", "Edge70", _
        "Edge71", "Edge72")
        
End Function

Private Function GET_ALL_INTERSECTIONS() As Variant

    GET_ALL_INTERSECTIONS = Array("int1", "int2", "int3", "int4", "int5", "int6", "int7", _
    "int8", "int9", "int10", "int11", "int12", "int13", "int14", "int15", "int16", "int17", _
    "int18", "int19", "int20", "int21", "int22", "int23", "int24", "int25", "int26", _
    "int27", "int28", "int29", "int30", "int31", "int32", "int33", "int34", "int35", _
    "int36", "int37", "int38", "int39", "int40", "int41", "int42", "int43", "int44", _
    "int45", "int46", "int47", "int48", "int49", "int50", "int51", "int52", "int53", _
    "int54")

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



Function reset_edge_intersections_trackers()
'
' reset_edge_intersections_trackers Macro
' resets tables rexsponsible for tracking edges and intersections
'

'
    Sheets("Sheet2").Select
    Range("edge_tracker[Road]").Select
    Selection.ClearContents
    Range("intersection_tracker[City/settlement]").Select
    Selection.ClearContents
    Range("board_tracker[Robber]").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E3").Select
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E3").Select
    Range("E3:E4").Select
    Selection.AutoFill Destination:=Range("board_tracker[Robber]"), Type:= _
        xlFillDefault
    Range("board_tracker[Robber]").Select
    Selection.AutoFill Destination:=Range("board_tracker[[Cities]:[Robber]]"), _
        Type:=xlFillDefault
    Range("board_tracker[[Cities]:[Robber]]").Select
    Sheets("Sheet1").Select
End Function



Public Function check_edges()
    Dim edges As ShapeRange
    Set edges = ActiveSheet.shapes.Range(GET_ALL_EDGES())
    
    Dim edge_tracker As Range
    Set edge_tracker = Sheets("Sheet2").Range("edge_tracker")
    
    Dim player_stats As Range
    Set player_stats = ActiveSheet.Range("Player_stats")
    
    Dim i As Long
    For i = 1 To edges.Count
        Dim edge_name As String
        edge_name = edges(i).Name
        Dim road As String
        road = edge_tracker.Cells(i, 2).Value
        
        Select Case road
                Case "Player 1"
                    edges(i).Line.ForeColor.RGB = RGB(0, 176, 240)
                    edges(i).Line.Weight = 5
                Case "Player 2"
                    edges(i).Line.ForeColor.RGB = RGB(102, 255, 51)
                    edges(i).Line.Weight = 5
                Case "Player 3"
                    edges(i).Line.ForeColor.RGB = RGB(255, 0, 0)
                    edges(i).Line.Weight = 5
                Case "Player 4"
                    edges(i).Line.ForeColor.RGB = RGB(255, 255, 0)
                    edges(i).Line.Weight = 5
                Case ""
                    edges(i).Line.ForeColor.RGB = RGB(0, 0, 0)
                    edges(i).Line.Weight = 2
        End Select
        
    Next i
    
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
                shapes(i).Fill.ForeColor.RGB = colorRange.Item(colorIndex).Interior.color
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

        Call reset_edge_intersections_trackers
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
    Call check_edges
    ActiveSheet.shapes.Range(GET_ALL_INTERSECTIONS()).Select
    With Selection.ShapeRange
        .Line.Weight = 0
    End With

End Sub



Public Sub Worksheet_BeforeDoubleClick(ByVal Target As shape, Cancel As Boolean)
    If Target.Type = msoLine Then
        Call Select_edge(Target)
    End If
End Sub



Sub Select_edge()
    'Select_edge Macro
    'when an edge is clicked, it changes to white to show it is selected
    '
    Dim selectedShape As String, sh As shape
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
    
    Dim selectedShape As String, sh As shape
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



Sub Build_Road()
    'Build_Road Macro
    'Builds a road on the board based on user input
    '
    Dim selected_edge As String
    selected_edge = Sheets("Sheet2").Range("J2").Value

    Dim edge_tracker As Range
    Set edge_tracker = Sheets("Sheet2").Range("edge_tracker")
    Dim road As String
    
    Dim colorRange As Range
    Set colorRange = Sheets("Sheet1").Range("L17:L20")

    Dim shape As shape
    
    Dim color As Long
    Dim i As Long
    For i = 1 To edge_tracker.Rows.Count
        If edge_tracker.Cells(i, 1).Value = selected_edge Then
            road = edge_tracker.Cells(i, 2).Value
            Exit For
        End If
    Next i

    If road = "" Then
        Dim player As String
        player = Sheets("Sheet2").Range("K2").Value
        Dim player_stats As Range
        Set player_stats = Sheets("Sheet1").Range("Player_stats")
        Dim brick As Long
        Dim lumber As Long

        Dim j As Long
        For j = 1 To player_stats.Rows.Count
            If player_stats.Cells(j, 1).Value = player Then
                brick = player_stats.Cells(j, 3).Value
                lumber = player_stats.Cells(j, 5).Value
                Exit For
            End If
        Next j

        If brick > 0 And lumber > 0 Then
            Dim response As Integer
            response = MsgBox("Build a road?", vbYesNo + vbQuestion, "Reset Game")
            If response = vbYes Then
                road = player
                edge_tracker.Cells(i, 2).Value = road
                color = colorRange.Find(road).Interior.color
                Set shape = Sheets("Sheet1").shapes(selected_edge)
                With shape.Line
                    .Visible = msoTrue
                    .ForeColor.RGB = color
                    .Weight = 5
                    .Transparency = 0
                End With
                player_stats.Cells(j, 3).Value = brick - 1
                player_stats.Cells(j, 5).Value = lumber - 1
            End If
        Else
            MsgBox "Not enough resources", vbOKOnly + vbExclamation, "Error"

        End If
    Else
        color = colorRange.Find(road).Interior.color
        Set shape = Sheets("Sheet1").shapes(selected_edge)
        With shape.Line
            .Visible = msoTrue
            .ForeColor.RGB = color
            .Weight = 5
            .Transparency = 0
        End With
        MsgBox "Edge already has a road", vbOKOnly + vbExclamation, "Error"
    End If
    
End Sub





Public Sub Build_Settlement()
    'Build_Settlement Macro
    'Builds a settlement on the board based on user input
    '
    Dim selected_intersection As String
    selected_intersection = Sheets("Sheet2").Range("J2").Value

    Dim intersection_tracker As Range
    Set intersection_tracker = Sheets("Sheet2").Range("intersection_tracker")

    Dim settlement As String

    Dim colorRange As Range
    Set colorRange = Sheets("Sheet1").Range("L17:L20")

    Dim shape As shape

    Dim settlement_shape As shape

    Dim color As Long
    Dim i As Long
    For i = 1 To intersection_tracker.Rows.Count
        If intersection_tracker.Cells(i, 1).Value = selected_intersection Then
            settlement = intersection_tracker.Cells(i, 2).Value
            Exit For
        End If
    Next i

    If settlement = "" Then
        Dim player As String
        player = Sheets("Sheet2").Range("K2").Value
        Dim player_stats As Range
        Set player_stats = Sheets("Sheet1").Range("Player_stats")
        Dim brick As Long
        Dim lumber As Long
        Dim wool As Long
        Dim wheat As Long

        Dim j As Long
        For j = 1 To player_stats.Rows.Count
            If player_stats.Cells(j, 1).Value = player Then
                wheat = player_stats.Cells(j, 2).Value
                brick = player_stats.Cells(j, 3).Value
                wool = player_stats.Cells(j, 4).Value
                lumber = player_stats.Cells(j, 5).Value
                Exit For
            End If
        Next j

        If brick > 0 And lumber > 0 And wool > 0 And wheat > 0 Then
            Dim response As Integer
            response = MsgBox("Build a settlement?", vbYesNo + vbQuestion, "Build Settlement")
            If response = vbYes Then
                settlement = player
                intersection_tracker.Cells(i, 2).Value = settlement
                color = colorRange.Find(settlement).Interior.color
                Set shape = Sheets("Sheet1").shapes(selected_intersection)
                Set settlement_shape = Sheets("Sheet1").shapes("Build_Settlement").Duplicate
                settlement_shape.Fill.ForeColor.RGB = color
                settlement_shape.Left = shape.Left
                settlement_shape.Top = shape.Top
                player_stats.Cells(j, 3).Value = brick - 1
                player_stats.Cells(j, 5).Value = lumber - 1
            End If
        Else
            MsgBox "Not enough resources", vbOKOnly + vbExclamation, "Error"

        End If
    Else
        MsgBox "Intersection already has a settlement", vbOKOnly + vbExclamation, "Error"
    End If
End Sub



Public Sub Build_City()
    'Build_City Macro
    'Builds a city on the board based on user input
    '
    Dim selected_intersection As String
    selected_intersection = Sheets("Sheet2").Range("J2").Value

    Dim intersection_tracker As Range
    Set intersection_tracker = Sheets("Sheet2").Range("intersection_tracker")

    Dim city As String

    Dim colorRange As Range
    Set colorRange = Sheets("Sheet1").Range("L17:L20")

    Dim shape As shape

    Dim city_shape As shape

    Dim color As Long
    Dim i As Long
    For i = 1 To intersection_tracker.Rows.Count
        If intersection_tracker.Cells(i, 1).Value = selected_intersection Then
            city = intersection_tracker.Cells(i, 2).Value
            Exit For
        End If
    Next i

    If city = Sheets("Sheet2").Range("K2").Value Then
        Dim player As String
        player = Sheets("Sheet2").Range("K2").Value
        Dim player_stats As Range
        Set player_stats = Sheets("Sheet1").Range("Player_stats")
        Dim wheat As Long
        Dim ore As Long

        Dim j As Long
        For j = 1 To player_stats.Rows.Count
            If player_stats.Cells(j, 1).Value = player Then
                wheat = player_stats.Cells(j, 2).Value
                ore = player_stats.Cells(j, 6).Value
                Exit For
            End If
        Next j

        If wheat > 2 And ore > 3 Then
            Dim response As Integer
            response = MsgBox("Build a city?", vbYesNo + vbQuestion, "Build City")
            If response = vbYes Then
                city = player
                intersection_tracker.Cells(i, 2).Value = city
                color = colorRange.Find(city).Interior.color
                Set shape = Sheets("Sheet1").shapes(selected_intersection)
                Set city_shape = Sheets("Sheet1").shapes("Build_City").Duplicate
                city_shape.Fill.ForeColor.RGB = color
                city_shape.Left = shape.Left
                city_shape.Top = shape.Top
                player_stats.Cells(j, 2).Value = wheat - 2
                player_stats.Cells(j, 6).Value = ore - 3
            End If
        Else
            MsgBox "Not enough resources", vbOKOnly + vbExclamation, "Error"
    
        End If
    Else
        MsgBox "You cannot build a settlement here", vbOKOnly + vbExclamation, "Error"
    End If
End Sub




