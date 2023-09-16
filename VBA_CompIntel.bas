Attribute VB_Name = "Module1"
Sub CreateTableOfContentsWithFormatting()
    Dim wsTOC As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tocRow As Long

    ' Add a new sheet and name it "Table of Contents"
    Set wsTOC = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1)) ' Move to the beginning
    wsTOC.Name = "Table of Contents"

    ' Format the "Table of Contents" sheet
    With wsTOC
        .Cells.ClearContents ' Clear existing contents
        .Range("A1").Value = "Table of Contents"
        .Range("B1").Value = "SecId" ' Add "SecId" beside the header
        .Range("C1").Value = "Peer Category" ' Add "SecId" beside the header
        .Range("A1:C1").Font.Size = 18 ' Increase font size
        .Range("A1:C1").Font.Bold = True
        .Columns("A:C").AutoFit
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 40
        .Cells.Font.Name = "Century Gothic"
        .Cells.Font.Size = 10
        .Range("A1:C1").Interior.Color = RGB(52, 73, 94) ' Dark blue background
        .Range("A1:C1").Font.Color = RGB(255, 255, 255) ' White font color
        .Range("A1:C1").HorizontalAlignment = xlCenter ' Center align header text
        .Rows("2:" & .Rows.Count).HorizontalAlignment = xlCenter ' Center align all rows starting from row 2
    End With

    tocRow = 2 ' Start listing the sheets from row 2

    ' Loop through all the sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Skip the "Table of Contents" sheet itself
        If ws.Name <> "Table of Contents" Then
        
            ' Freeze rows 1 and 2 in the current sheet
            ws.Activate
            ActiveWindow.FreezePanes = False ' Unfreeze existing panes
            ws.Range("A3").Select ' Select a cell below the frozen rows
            ActiveWindow.FreezePanes = True ' Freeze rows 1 and 2
        
            ' Autofit all columns in the current sheet
            ws.Cells.EntireColumn.AutoFit

            ' Set font to Century Gothic, size 11 in the current sheet
            ws.Cells.Font.Name = "Century Gothic"
            ws.Cells.Font.Size = 10

            ' Center align all cells in the current sheet
            ws.Cells.HorizontalAlignment = xlCenter

            ' Bold all cells in row 2 of the non-table of contents sheets
            ws.Rows(2).Font.Bold = True

            ' Add the sheet name and cell C2 value to the "Table of Contents" sheet
            wsTOC.Cells(tocRow, 1).Value = IIf(Not IsEmpty(ws.Range("C2").Value), ws.Range("C2").Value, ws.Name)
            wsTOC.Cells(tocRow, 3).Value = IIf(Not IsEmpty(ws.Range("D2").Value), ws.Range("D2").Value, ws.Name)
            wsTOC.Cells(tocRow, 2).Value = ws.Name

            ' Add hyperlink to the sheet name
            With wsTOC.Hyperlinks.Add(Anchor:=wsTOC.Cells(tocRow, 1), _
                                      Address:="", _
                                      SubAddress:="'" & ws.Name & "'!A1", _
                                      TextToDisplay:=IIf(Not IsEmpty(ws.Range("C2").Value), ws.Range("C2").Value, ws.Name))
                .ScreenTip = ws.Name
            End With

            ' Format the cells in the "Table of Contents" sheet
            wsTOC.Cells(tocRow, 1).Font.Underline = xlUnderlineStyleSingle
            wsTOC.Cells(tocRow, 1).Font.Color = RGB(0, 0, 255) ' Blue font color

            tocRow = tocRow + 1

            ' Apply custom formatting to columns F and G
            Dim lastDataRow As Long
            lastDataRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

            With ws.Range("F2:G" & lastDataRow)
                .Borders.LineStyle = xlNone ' Remove border lines for columns F and G
                .FormatConditions.Delete ' Clear previous conditional formatting
            End With

            ' Apply custom formatting to columns I and J
            With ws.Range("I2:J" & lastDataRow)
                .Borders.LineStyle = xlNone ' Remove border lines for columns I and J
                .FormatConditions.Delete ' Clear previous conditional formatting
            End With

            ' Apply conditional formatting for color bars in columns F and G
            Dim colF As Range
            Dim colG As Range
            Dim maxValueFG As Double
            Dim minValueFG As Double

            Set colF = ws.Range("F2:F" & lastDataRow)
            Set colG = ws.Range("G2:G" & lastDataRow)

            maxValueF = Application.WorksheetFunction.Max(colF)
            minValueF = Application.WorksheetFunction.Min(colF)
            maxValueG = Application.WorksheetFunction.Max(colG)
            minValueG = Application.WorksheetFunction.Min(colG)

            With colF.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).Value = minValueF
                .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) ' Red
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).Value = (maxValueF - minValueF) / 2
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0) ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueNumber
                .ColorScaleCriteria(3).Value = maxValueF
                .ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0) ' Green
            End With
            
            With colG.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).Value = minValueG
                .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) ' Red
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).Value = (maxValueG - minValueG) / 2
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0) ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueNumber
                .ColorScaleCriteria(3).Value = maxValueG
                .ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0) ' Green
            End With
            

            ' Apply conditional formatting for color bars in columns I and J
            Dim colI As Range
            Dim colJ As Range
            Dim maxValueIJ As Double
            Dim minValueIJ As Double

            Set colI = ws.Range("I2:I" & lastDataRow)
            Set colJ = ws.Range("J2:J" & lastDataRow)

            maxValueI = Application.WorksheetFunction.Max(colI)
            minValueI = Application.WorksheetFunction.Min(colI)
            maxValueJ = Application.WorksheetFunction.Max(colJ)
            minValueJ = Application.WorksheetFunction.Min(colJ)

            With colI.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).Value = minValueI
                .ColorScaleCriteria(1).FormatColor.Color = RGB(0, 255, 0) ' Green
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).Value = (maxValueI - minValueI) / 2
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0) ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueNumber
                .ColorScaleCriteria(3).Value = maxValueI
                .ColorScaleCriteria(3).FormatColor.Color = RGB(255, 0, 0) ' Red
            End With

            With colJ.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).Value = minValueJ
                .ColorScaleCriteria(1).FormatColor.Color = RGB(0, 255, 0) ' Green
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).Value = (maxValueJ - minValueJ) / 2
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0) ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueNumber
                .ColorScaleCriteria(3).Value = maxValueJ
                .ColorScaleCriteria(3).FormatColor.Color = RGB(255, 0, 0) ' Red
            End With

            ' Apply conditional formatting for color bars in columns K and L
            Dim colK As Range
            Dim colL As Range
            Dim maxValueKL As Double
            Dim minValueKL As Double

            Set colK = ws.Range("K2:K" & lastDataRow)
            Set colL = ws.Range("L2:L" & lastDataRow)

            maxValueK = Application.WorksheetFunction.Max(colK)
            minValueK = Application.WorksheetFunction.Min(colK)
            maxValueL = Application.WorksheetFunction.Max(colL)
            minValueL = Application.WorksheetFunction.Min(colL)
            
            With colK.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).Value = minValueK
                .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) ' Red
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).Value = (maxValueK - minValueK) / 2
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0) ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueNumber
                .ColorScaleCriteria(3).Value = maxValueK
                .ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0) ' Green
            End With
            
            With colL.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).Value = minValueL
                .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) ' Red
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).Value = (maxValueL - minValueL) / 2
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0) ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueNumber
                .ColorScaleCriteria(3).Value = maxValueL
                .ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0) ' Green
            End With

            ' Apply conditional formatting for color bars in columns M
            Dim colM As Range

            Set colM = ws.Range("M2:M" & lastDataRow)

            maxValueM = Application.WorksheetFunction.Max(colM)
            minValueM = Application.WorksheetFunction.Min(colM)

            With colM.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).Value = minValueM
                .ColorScaleCriteria(1).FormatColor.Color = RGB(0, 255, 0) ' Green
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).Value = (maxValueM - minValueM) / 2
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0) ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueNumber
                .ColorScaleCriteria(3).Value = maxValueM
                .ColorScaleCriteria(3).FormatColor.Color = RGB(255, 0, 0) ' Red
            End With
            
        End If
    Next ws

    ' AutoFit the columns after adding all the sheet names
    wsTOC.Columns("A:B").AutoFit

    ' Scroll back to the top of the "Table of Contents" sheet
    wsTOC.Activate
    wsTOC.Range("A1").Select

    ' Remove borders from the "Table of Contents" sheet
    wsTOC.Cells.Borders.LineStyle = xlNone

    ' Add filters to rows E to K in all sheets except "Table of Contents"
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Table of Contents" Then
            ws.Range("E:M").AutoFilter
        End If
    Next ws
End Sub


