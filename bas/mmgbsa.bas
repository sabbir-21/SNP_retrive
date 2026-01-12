Sub Create_Thermal_MMGBSA_Chart_With_Textboxes()

    Dim ws As Worksheet
    Dim cht As Chart
    Dim co As ChartObject
    Dim rng As Range
    Dim s As Series
    Dim shp As Shape
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    'Set rng = ws.Range("A1:H4")
	Dim lastRow As Long, lastCol As Long
	lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
	lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
	Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Delete old charts
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    
    ' Delete old MMGBSA textboxes
    For Each shp In ws.Shapes
        If shp.Name Like "MMGBSA_Label_*" Then shp.Delete
    Next shp
    
    ' Create chart
    Set co = ws.ChartObjects.Add(Left:=350, Top:=20, Width:=700, Height:=400)
    Set cht = co.Chart
    
    With cht
        .SetSourceData Source:=rng
        .ChartType = xlBarClustered
        
        ' Title
        .HasTitle = True
        .ChartTitle.Text = "Thermal MMGBSA"
        With .ChartTitle.Font
            .Name = "Times New Roman"
            .Size = 14
            .Bold = True
        End With
        
        ' Legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Font.Name = "Times New Roman"
        
        ' REMOVE category axis (Y)
        With .Axes(xlCategory)
            .TickLabelPosition = xlNone
            .HasTitle = False
            .Format.Line.Visible = msoFalse
        End With
        
        ' Value axis (X)
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
        
        ' Data labels inside bars (2 decimals)
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
    Set terms = ws.Range("B1:H1")
    
    Dim plotTop As Double, plotHeight As Double
    plotTop = co.Top + cht.PlotArea.Top
    plotHeight = cht.PlotArea.Height
    
    For i = 1 To terms.Columns.Count
        
        Set shp = ws.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=co.Left - 170, _
            Top:=plotTop + (plotHeight / terms.Columns.Count) * (i - 0.5), _
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
