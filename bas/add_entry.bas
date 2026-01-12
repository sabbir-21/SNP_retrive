Attribute VB_Name = "add_db"
Sub SaveFormData()
    Dim wsForm As Worksheet, wsData As Worksheet
    Dim nextRow As Long
	Dim chkCell As Range
	
	' Path to the external file where you want to save the data
    dataFilePath = "D:\Research\Thesis\sample collection\Data.xlsx"
    
    Set wsForm = ThisWorkbook.Sheets("form")
	' Try to open the external workbook (if not already open)
    On Error Resume Next
    Set wbData = Workbooks.Open(dataFilePath)
    On Error GoTo 0

    ' If file not found, show message
    If wbData Is Nothing Then
        MsgBox "Data file not found at: " & dataFilePath, vbExclamation
        Exit Sub
    End If

    ' Reference to the "Data" sheet in that workbook
    Set wsData = wbData.Sheets("data")
    'Set wsData = ThisWorkbook.Sheets("data")
    
    ' Find next empty row in data sheet
    nextRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
	
	'------------------------------
    ' Convert checkbox TRUE/FALSE to Yes/No
    '------------------------------
    For Each chkCell In wsForm.Range("B12:D12,B13:D13,C14:D16,B19:D19")
        If chkCell.Value = True Then
            chkCell.Value = "Yes"
        ElseIf chkCell.Value = False Then
            chkCell.Value = "No"
        End If
    Next chkCell
    
    ' Transfer values from form to data
    wsData.Range("A" & nextRow).Value = wsForm.Range("B2").Value   ' Date
    wsData.Range("B" & nextRow).Value = wsForm.Range("D2").Value   ' Sample No
    wsData.Range("C" & nextRow).Value = wsForm.Range("B4").Value   ' Name
    wsData.Range("D" & nextRow).Value = wsForm.Range("D4").Value   ' Working Type
    wsData.Range("E" & nextRow).Value = wsForm.Range("B5").Value   ' Age
    wsData.Range("F" & nextRow).Value = wsForm.Range("D5").Value   ' Gender
    wsData.Range("G" & nextRow).Value = wsForm.Range("B6").Value   ' Occupation
    wsData.Range("H" & nextRow).Value = wsForm.Range("D6").Value   ' Contact No
    wsData.Range("I" & nextRow).Value = wsForm.Range("B7").Value   ' Residence
    wsData.Range("J" & nextRow).Value = wsForm.Range("D7").Value  ' Height
    wsData.Range("K" & nextRow).Value = wsForm.Range("B8").Value  ' Weight
	wsData.Range("L" & nextRow).Value = wsForm.Range("D8").Value  ' BMI
    
    ' Patients' history & diagnostic
    wsData.Range("M" & nextRow).Value = wsForm.Range("B10").Value ' Type of Diabetes
    wsData.Range("N" & nextRow).Value = wsForm.Range("B11").Value ' Date when Diagnose
    wsData.Range("O" & nextRow).Value = wsForm.Range("B12").Value ' Diet 
	wsData.Range("P" & nextRow).Value = wsForm.Range("C12").Value ' Oral Meds
	wsData.Range("Q" & nextRow).Value = wsForm.Range("D12").Value ' Insulin
	
    wsData.Range("R" & nextRow).Value = wsForm.Range("B13").Value ' Nephropathy
	wsData.Range("S" & nextRow).Value = wsForm.Range("C13").Value ' Retinopathy
	wsData.Range("T" & nextRow).Value = wsForm.Range("D13").Value ' Neuropathy
	wsData.Range("U" & nextRow).Value = wsForm.Range("C14").Value ' Heart Disease
	wsData.Range("V" & nextRow).Value = wsForm.Range("D14").Value ' Hypertension
	wsData.Range("W" & nextRow).Value = wsForm.Range("C15").Value ' Stroke
	wsData.Range("X" & nextRow).Value = wsForm.Range("D15").Value ' Liver Disease
	wsData.Range("Y" & nextRow).Value = wsForm.Range("C16").Value ' Kidney Disease 
	wsData.Range("Z" & nextRow).Value = wsForm.Range("D16").Value ' RA
	
    wsData.Range("AA" & nextRow).Value = wsForm.Range("B14").Value ' Family History of Diabetes
    wsData.Range("AB" & nextRow).Value = wsForm.Range("B15").Value ' Regular balanced Diet
    wsData.Range("AC" & nextRow).Value = wsForm.Range("B16").Value ' Exercise Regularly
    wsData.Range("AD" & nextRow).Value = wsForm.Range("B17").Value ' Smoking Habits
	wsData.Range("AE" & nextRow).Value = wsForm.Range("B18").Value ' Sugary Drinks
    
	wsData.Range("AF" & nextRow).Value = wsForm.Range("B19").Value ' Fish
	wsData.Range("AG" & nextRow).Value = wsForm.Range("C19").Value ' Meat
	wsData.Range("AH" & nextRow).Value = wsForm.Range("D19").Value ' Vegetable
	
    wsData.Range("AI" & nextRow).Value = wsForm.Range("B21").Value ' Fasting Blood Sugar
    wsData.Range("AJ" & nextRow).Value = wsForm.Range("B22").Value ' After Breakfast Sugar
    wsData.Range("AK" & nextRow).Value = wsForm.Range("B23").Value ' HbA1C
    wsData.Range("AL" & nextRow).Value = wsForm.Range("B24").Value ' Total Cholesterol
    wsData.Range("AM" & nextRow).Value = wsForm.Range("B25").Value ' Blood Pressure
    wsData.Range("AN" & nextRow).Value = wsForm.Range("B26").Value ' HDL-C 
    wsData.Range("AO" & nextRow).Value = wsForm.Range("B27").Value ' LDL-C 
    wsData.Range("AP" & nextRow).Value = wsForm.Range("B28").Value ' eGFR
	wsData.Range("AQ" & nextRow).Value = wsForm.Range("B29").Value ' Creatinine
	wsData.Range("AR" & nextRow).Value = wsForm.Range("B30").Value ' Triglycerides
    
    MsgBox "Data Added Successfully!", vbInformation
    
    '------------------------------
    ' Clear form after saving
    '------------------------------
    With wsForm
        .Range("B4:B8").ClearContents
        .Range("D4:D7").ClearContents
        .Range("B10:B19").ClearContents
        .Range("B21:B30").ClearContents
		
		.Range("B12:D12").Value = False
		.Range("B13:D13").Value = False
		.Range("C14:D16").Value = False
		.Range("B19:D19").Value = False
    End With
    
    ' Refresh Date
    'wsForm.Range("B2").Value = Date
    
    ' Auto-generate next Sample No (SH001, SH002, etc.)
    wsForm.Range("D2").Value = "SH" & Format(nextRow, "000")
    
End Sub
