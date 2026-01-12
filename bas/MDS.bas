Attribute VB_Name = "Module1"
Sub CreateRMSDScatterPlot()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim series As Series
    Dim lastRow As Long, lastCol As Long
    Dim i As Integer
    Dim maxX As Double
    Dim colors As Object
    Dim colLetter As String
    Dim legendSeries As Series
    Dim axisTitle As String

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of data in column A (X-axis values)
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Find the last used column (all columns after A)
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' Get the maximum X-axis value from column A (Time values)
        maxX = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow))

        ' Delete existing charts in the sheet
        For Each chartObj In ws.ChartObjects
            chartObj.Delete
        Next chartObj

        ' Add a new chart object
        Set chartObj = ws.ChartObjects.Add(Left:=500, Width:=400, Top:=50, Height:=266)
        Set chart = chartObj.Chart
        chart.ChartType = xlXYScatterLines ' Scatter plot with lines

        ' Define fixed colors for specific columns
        Set colors = CreateObject("Scripting.Dictionary")
        colors.Add "B", RGB(0, 0, 235) ' Blue
        colors.Add "C", RGB(255, 0, 0) ' Red
        colors.Add "D", RGB(0, 255, 0) ' Green
        colors.Add "E", RGB(255, 206, 86) ' Yellow
        colors.Add "F", RGB(153, 102, 255) ' Purple
        colors.Add "G", RGB(255, 159, 64) ' Orange
        colors.Add "H", RGB(54, 162, 140) ' Teal
        colors.Add "I", RGB(201, 203, 207) ' Gray

        ' Loop through all columns after A to add series
        For i = 2 To lastCol
            Set series = chart.SeriesCollection.NewSeries
            series.Name = ws.Cells(1, i).Value                           ' Series name from row 1
            series.XValues = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)) ' X-values from column A
            series.Values = ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i))  ' Y-values from column i

            ' Get column letter and set color based on mapping
            colLetter = Split(ws.Cells(1, i).Address, "$")(1)
            If colors.exists(colLetter) Then
                series.Border.Color = colors(colLetter)
            Else
                series.Border.Color = vbBlack ' Default color if not mapped
            End If

            series.MarkerStyle = xlNone  ' Remove markers for smooth lines
        Next i

        ' Set dynamic Y-axis title based on sheet name
        Select Case ws.Name
            Case "RMSD"
                axisTitle = "RMSD (" & ChrW(&H00C5) & ")"
			Case "RMSD_Protein"
                axisTitle = "Protein RMSD (" & ChrW(&H00C5) & ")"
			Case "RMSD_LIG"
                axisTitle = "Ligand RMSD (" & ChrW(&H00C5) & ")"
            Case "RMSF"
                axisTitle = "RMSF (" & ChrW(&H00C5) & ")"
            Case "RG"
                axisTitle = "Radius of Gyration (" & ChrW(&H00C5) & ")"
            Case "SASA"
                axisTitle = "SASA (" & ChrW(&H00C5) & ChrW(178) & ")"
			Case "PSA"
                axisTitle = "PSA (" & ChrW(&H00C5) & ChrW(178) & ")"
			Case "MOLSA"
                axisTitle = "MOLSA (" & ChrW(&H00C5) & ChrW(178) & ")"
			Case "HB"
                axisTitle = "Hydrogen Bonds"
            Case Else
                axisTitle = ws.Name ' Default to sheet name if not listed
        End Select

        ' Format the chart
        With chart
            .HasTitle = False
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = ws.Range("A1").Value
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = axisTitle
            .Legend.Position = xlBottom
            .HasLegend = True
            .PlotArea.Format.Line.Visible = msoFalse ' Remove border around plot area
            .ChartArea.Format.Line.Visible = msoFalse ' Remove border around chart area
            .Axes(xlCategory).MajorGridlines.Delete ' Remove X-axis gridlines
            .Axes(xlValue).MajorGridlines.Delete ' Remove Y-axis gridlines
        End With

		' Set X-axis scale dynamically
		With chart.Axes(xlCategory)
			.MinimumScale = 0
			.MajorUnit = Application.WorksheetFunction.Round(((maxX \ 10) + 1), -1) ' Round to nearest 10
			.MaximumScale = maxX
		End With

        ' Adjust line thickness
        For Each legendSeries In chart.SeriesCollection
            legendSeries.Format.Line.Weight = 2.25 ' Adjust thickness (increase for thicker lines)
        Next legendSeries

        ' Apply font settings
        With chart.Axes(xlCategory)
            .TickLabels.Font.Name = "Times New Roman"
            .TickLabels.Font.Size = 14
            .TickLabels.Font.Bold = True
            .AxisTitle.Font.Name = "Times New Roman"
            .AxisTitle.Font.Size = 14
            .AxisTitle.Font.Bold = True
        End With

        With chart.Axes(xlValue)
            .TickLabels.Font.Name = "Times New Roman"
            .TickLabels.Font.Size = 14
            .TickLabels.Font.Bold = True
            .AxisTitle.Font.Name = "Times New Roman"
            .AxisTitle.Font.Size = 14
            .AxisTitle.Font.Bold = True
        End With

        With chart.Legend
            .Font.Name = "Times New Roman"
            .Font.Size = 14
            .Font.Bold = True
        End With

        ' Cleanup
        Set colors = Nothing
    Next ws ' Move to the next sheet

End Sub
