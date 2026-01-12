Sub Create_Thermal_MMGBSA_Chart_With_Textboxes()

    Dim ws As Worksheet
    Dim cht As Chart
    Dim co As ChartObject
    Dim rng As Range
    Dim s As Series
    Dim shp As Shape
    Dim i As Long
    Dim lastRow As Long, lastCol As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' -------- DYNAMIC DATA RANGE --------
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' -------- DELETE OLD CHARTS --------
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    
    ' -------- DELETE OLD TEXTBOXES --------
    For Each shp In ws.Shapes
        If shp.Name Like "MMGBSA_Label_*" Then shp.Delete
    Next shp
    
    ' -------- CREATE CHART --------
    Set co = ws.ChartObjects.Add(Left:=350, Top:=20, Width:=700, Height:=400)
    Set cht = co.Chart
    
    With cht
        .SetSourceData Source:=rng
        .ChartType = xlBarClustered
        
        ' ----- TITLE -----
        .HasTitle = True
        .ChartTitle.Text = "Thermal MMGBSA"
        With .ChartTitle.Font
            .Name = "Times New Roman"
            .Size = 14
            .Bold = True
        End With
        
        ' ----- LEGEND -----
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Font.Name = "Times New Roman"
        
        ' ----- REMOVE CATEGORY AXIS -----
        With .Axes(xlCategory)
            .TickLabelPosition = xlNone
            .HasTitle = False
            .Format.Line.Visible = msoFalse
        End With
        
        ' ----- VALUE AXIS -----
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Average Energy (kcal/mol)"
            With .AxisTitle.Font
                .Name = "Times New Roman"
                .Size = 11
            End With
            .CrossesAt = 0
            .TickLabels.Font.Size = 9
            .HasMajorGridlines = False
        End With
        
        ' ----- DATA LABELS -----
        For Each s In .SeriesCollection
            s.HasDataLabels = True
            With s.DataLabels
                .ShowValue = True
                .Position = xlLabelPositionInsideEnd
                .NumberFormat = "0.00"
                .Font.Name = "Times New Roman"
                .Font.Size = 8
            End With
        Next s
        
        .ChartGroups(1).GapWidth = 60
    End With
    
    ' -------- ADD MMGBSA TERM TEXTBOXES --------
    
    Dim terms As Range
    Set terms = ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol))   ' B1 → lastCol
    
    Dim plotTop As Double, plotHeight As Double
    plotTop = co.Top + cht.PlotArea.Top
    plotHeight = cht.PlotArea.Height
    
    ' IMPORTANT: reverse vertical placement to match bar order
    For i = 1 To terms.Columns.Count
        
        Set shp = ws.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=co.Left - 170, _
            Top:=plotTop + (plotHeight / terms.Columns.Count) * (terms.Columns.Count - i + 0.5), _
            Width:=170, _
            Height:=18 _
        )
        
        With shp
            .Name = "MMGBSA_Label_" & i
            .TextFrame.Characters.Text = terms.Cells(1, i).Value
            .TextFrame.HorizontalAlignment = xlHAlignRight
            .TextFrame.VerticalAlignment = xlVAlignCenter
            
            With .TextFrame.Characters.Font
                .Name = "Times New Roman"
                .Size = 9
                .Bold = True
            End With
            
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
        End With
        
    Next i

End Sub
