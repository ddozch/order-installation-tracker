Attribute VB_Name = "modDashboard"
'==============================================================================
' Installation Tracker
' Dashboard layout. Every figure is a dynamic-array formula over the tables, so
' changing Year or Product recalculates instantly without running a macro.
' Formulas are written with Formula2 in English; Excel shows them in its own language.
'==============================================================================
Option Explicit

' Product filter: "All" matches every product through the * wildcard
Private Const LET_PROD As String = "prod,IF(selProduct=""All"",""*"",selProduct)"
' Year to date: from 1 January of the selected year to the month of the report date
Private Const LET_YTD As String = "yStart,DATE(selYear,1,1),mEnd,DATE(YEAR(cfgReportDate),MONTH(cfgReportDate),1)"
' Quarter start and the start of the next quarter
Private Const LET_QTR As String = "qStart,DATE(selYear,SEQUENCE(4,1,1,3),1),qEnd,DATE(selYear,SEQUENCE(4,1,4,3),1)"

Public Sub BuildDashboard(ByVal ws As Worksheet)
    Dim i As Long
    ws.Cells.Clear
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C:J").ColumnWidth = 11
    ws.Columns("K").ColumnWidth = 3
    ws.Columns("L:M").ColumnWidth = 11

    With ws.Range("B1")
        .Value = "Installation Tracker"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = ClrAccent()
    End With
    With ws.Range("B2")
        .Value = "Not updated yet. Click Update."
        .Font.Color = ClrMuted()
    End With
    AddName NM_LAST_UPDATE, ws.Range("B2")

    BuildSettings ws
    BuildKpis ws
    BuildMonthly ws
    BuildQuarters ws
    BuildProducts ws
    BuildChart ws
End Sub

Private Sub BuildSettings(ByVal ws As Worksheet)
    Dim shp As Shape
    SettingRow ws, 4, "Report date", DateSerial(2026, 10, 30), "yyyy-mm-dd", NM_REPORT_DATE
    SettingRow ws, 5, "Year", 2026, "0", NM_YEAR
    SettingRow ws, 6, "Product", "All", "@", NM_PRODUCT
    SettingRow ws, 7, "Data folder", "data", "@", NM_DATA_FOLDER
    With ws.Range("C6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & SH_LISTS & "!$A$1:$A$12"
    End With

    Set shp = ws.Shapes.AddShape(5, ws.Range("E4").Left, ws.Range("E4").Top, _
                                 ws.Range("E4:G4").Width, ws.Range("E4:E6").Height)   ' 5 = rounded rectangle
    With shp
        .Name = "btnUpdate"
        .OnAction = "UpdateTracker"
        .Fill.ForeColor.RGB = ClrAccent()
        .Line.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        With .TextFrame2.TextRange
            .Text = "Update"
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Size = 16
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    With ws.Range("H4")
        .Value = "Reads orders, crew schedules, contracts and specifications, then runs every check."
        .Font.Color = ClrMuted()
        .Font.Italic = True
    End With
End Sub

Private Sub BuildKpis(ByVal ws As Worksheet)
    ws.Range("B9:M10").Interior.Color = ClrAccentLight()
    Kpi ws, "B", "Plan YTD, units", _
        "=LET(" & LET_PROD & "," & LET_YTD & ",SUMIFS(tblOrders[Planned Units],tblOrders[Product],prod," & _
        "tblOrders[Planned Month],"">=""&yStart,tblOrders[Planned Month],""<=""&mEnd))", "#,##0"
    Kpi ws, "D", "Installed YTD, units", _
        "=LET(" & LET_PROD & "," & LET_YTD & ",SUMIFS(tblOrders[Installed Units],tblOrders[Product],prod," & _
        "tblOrders[Finish Month],"">=""&yStart,tblOrders[Finish Month],""<=""&mEnd))", "#,##0"
    Kpi ws, "F", "% of plan YTD", "=IF(B10=0,"""",D10/B10)", "0%"
    Kpi ws, "H", "Open orders YTD", _
        "=LET(" & LET_PROD & "," & LET_YTD & ",COUNTIFS(tblOrders[Product],prod,tblOrders[Planned Month],"">=""&yStart," & _
        "tblOrders[Planned Month],""<=""&mEnd,tblOrders[Status],""<>" & ST_INSTALLED & """))", "0"
    Kpi ws, "J", "Delayed start", _
        "=LET(" & LET_PROD & ",COUNTIFS(tblOrders[Product],prod,tblOrders[Status],""" & ST_DELAYED & """))", "0"
    Kpi ws, "L", "Open issues", _
        "=COUNTA(tblScheduleIssues[Key])+COUNTA(tblSpecIssues[Key])+COUNTA(tblContractIssues[Key])", "0"
End Sub

Private Sub BuildMonthly(ByVal ws As Worksheet)
    SectionTitle ws.Range("B12:J12"), "Plan vs actual by month"
    HeaderRow ws.Range("B13:J13"), Array("Month", "Plan", "Sold", "Installed", "% of plan", "Specified", _
                                         "Delayed", "Open", "No manager")
    ws.Range("B14").Formula2 = "=DATE(selYear,SEQUENCE(12),1)"
    ws.Range("C14").Formula2 = "=LET(" & LET_PROD & ",SUMIFS(tblOrders[Planned Units],tblOrders[Planned Month],B14#,tblOrders[Product],prod))"
    ws.Range("D14").Formula2 = "=LET(" & LET_PROD & ",SUMIFS(tblOrders[Planned Units],tblOrders[Order Month],B14#,tblOrders[Product],prod))"
    ws.Range("E14").Formula2 = "=LET(" & LET_PROD & ",SUMIFS(tblOrders[Installed Units],tblOrders[Finish Month],B14#,tblOrders[Product],prod))"
    ' No percentage for months after the report date: they cannot be complete yet
    ws.Range("F14").Formula2 = "=IF((C14#=0)+(B14#>cfgReportDate),"""",E14#/C14#)"
    ws.Range("G14").Formula2 = "=LET(" & LET_PROD & ",SUMIFS(tblSpecs[Units],tblSpecs[Spec Month],B14#,tblSpecs[Product],prod))"
    ws.Range("H14").Formula2 = "=LET(" & LET_PROD & ",COUNTIFS(tblOrders[Planned Month],B14#,tblOrders[Status],""" & _
                               ST_DELAYED & """,tblOrders[Product],prod))"
    ws.Range("I14").Formula2 = "=LET(" & LET_PROD & ",COUNTIFS(tblOrders[Planned Month],B14#,tblOrders[Status],""<>" & _
                               ST_INSTALLED & """,tblOrders[Product],prod))"
    ws.Range("J14").Formula2 = "=LET(" & LET_PROD & ",COUNTIFS(tblOrders[Planned Month],B14#,tblOrders[Manager],"""",tblOrders[Product],prod))"
    ws.Range("B14:B25").NumberFormat = "[$-409]mmm yyyy"
    ws.Range("B14:B25").HorizontalAlignment = xlLeft
    ws.Range("C14:E25,G14:J25").NumberFormat = "#,##0"
    ws.Range("F14:F25").NumberFormat = "0%"
    WarnBelowTarget ws.Range("F14:F25")
End Sub

Private Sub BuildQuarters(ByVal ws As Worksheet)
    SectionTitle ws.Range("B27:J27"), "By quarter"
    HeaderRow ws.Range("B28:G28"), Array("Quarter", "Plan", "Sold", "Installed", "% of plan", "Specified")
    ws.Range("B29").Formula2 = "=""Q""&SEQUENCE(4)"
    ws.Range("C29").Formula2 = "=LET(" & LET_PROD & "," & LET_QTR & ",SUMIFS(tblOrders[Planned Units]," & _
                               "tblOrders[Planned Month],"">=""&qStart,tblOrders[Planned Month],""<""&qEnd,tblOrders[Product],prod))"
    ws.Range("D29").Formula2 = "=LET(" & LET_PROD & "," & LET_QTR & ",SUMIFS(tblOrders[Planned Units]," & _
                               "tblOrders[Order Month],"">=""&qStart,tblOrders[Order Month],""<""&qEnd,tblOrders[Product],prod))"
    ws.Range("E29").Formula2 = "=LET(" & LET_PROD & "," & LET_QTR & ",SUMIFS(tblOrders[Installed Units]," & _
                               "tblOrders[Finish Month],"">=""&qStart,tblOrders[Finish Month],""<""&qEnd,tblOrders[Product],prod))"
    ws.Range("F29").Formula2 = "=LET(" & LET_QTR & ",IF((C29#=0)+(qStart>cfgReportDate),"""",E29#/C29#))"
    ws.Range("G29").Formula2 = "=LET(" & LET_PROD & "," & LET_QTR & ",SUMIFS(tblSpecs[Units]," & _
                               "tblSpecs[Spec Month],"">=""&qStart,tblSpecs[Spec Month],""<""&qEnd,tblSpecs[Product],prod))"
    ws.Range("C29:E32,G29:G32").NumberFormat = "#,##0"
    ws.Range("F29:F32").NumberFormat = "0%"
    WarnBelowTarget ws.Range("F29:F32")
End Sub

Private Sub BuildProducts(ByVal ws As Worksheet)
    SectionTitle ws.Range("B34:J34"), "By product, year to date (all products)"
    HeaderRow ws.Range("B35:G35"), Array("Product", "Plan", "Installed", "% of plan", "Open orders", "Delayed")
    ws.Range("B36").Formula2 = "=SORT(UNIQUE(tblOrders[Product]))"
    ws.Range("C36").Formula2 = "=LET(" & LET_YTD & ",SUMIFS(tblOrders[Planned Units],tblOrders[Product],B36#," & _
                               "tblOrders[Planned Month],"">=""&yStart,tblOrders[Planned Month],""<=""&mEnd))"
    ws.Range("D36").Formula2 = "=LET(" & LET_YTD & ",SUMIFS(tblOrders[Installed Units],tblOrders[Product],B36#," & _
                               "tblOrders[Finish Month],"">=""&yStart,tblOrders[Finish Month],""<=""&mEnd))"
    ws.Range("E36").Formula2 = "=IF(C36#=0,"""",D36#/C36#)"
    ws.Range("F36").Formula2 = "=LET(" & LET_YTD & ",COUNTIFS(tblOrders[Product],B36#,tblOrders[Planned Month],"">=""&yStart," & _
                               "tblOrders[Planned Month],""<=""&mEnd,tblOrders[Status],""<>" & ST_INSTALLED & """))"
    ws.Range("G36").Formula2 = "=COUNTIFS(tblOrders[Product],B36#,tblOrders[Status],""" & ST_DELAYED & """)"
    ws.Range("C36:D45,F36:G45").NumberFormat = "#,##0"
    ws.Range("E36:E45").NumberFormat = "0%"
    WarnBelowTarget ws.Range("E36:E45")
End Sub

Private Sub BuildChart(ByVal ws As Worksheet)
    Dim shp As Shape
    Set shp = ws.Shapes.AddChart2(-1, xlColumnClustered, ws.Range("L12").Left, ws.Range("L12").Top, 560, 330)
    shp.Name = "chtPlanActual"
    With shp.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Name = "Plan"
            .Values = ws.Range("C14:C25")
            .XValues = ws.Range("B14:B25")
            .Format.Fill.ForeColor.RGB = ClrPlan()
        End With
        With .SeriesCollection.NewSeries
            .Name = "Installed"
            .Values = ws.Range("E14:E25")
            .XValues = ws.Range("B14:B25")
            .Format.Fill.ForeColor.RGB = ClrAccent()
        End With
        .HasTitle = True
        .ChartTitle.Text = "Plan vs installed units by month"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlCategory).CategoryType = xlCategoryScale
        .Axes(xlCategory).TickLabels.NumberFormatLinked = True
        ' Linked to the cells: a format string on the axis is read in the Windows locale
        .Axes(xlValue).TickLabels.NumberFormatLinked = True
    End With
End Sub

Private Sub SettingRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal value As Variant, _
                       ByVal fmt As String, ByVal rangeName As String)
    ws.Cells(r, 2).Value = label
    With ws.Cells(r, 3)
        .NumberFormat = fmt
        .Value = value
        .Interior.Color = ClrInput()
        .HorizontalAlignment = xlLeft
    End With
    AddName rangeName, ws.Cells(r, 3)
End Sub

Private Sub Kpi(ByVal ws As Worksheet, ByVal col As String, ByVal label As String, ByVal formula As String, ByVal fmt As String)
    With ws.Range(col & "9")
        .Value = label
        .Font.Color = ClrMuted()
    End With
    With ws.Range(col & "10")
        .Formula2 = formula
        .NumberFormat = fmt
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

' Red bold text below 90%. "=9/10" avoids a decimal separator, which the rule would read in the local language.
Private Sub WarnBelowTarget(ByVal rng As Range)
    Dim fc As FormatCondition
    rng.FormatConditions.Delete
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=9/10")
    fc.Font.Color = ClrWarnText()
    fc.Font.Bold = True
End Sub

Public Sub SectionTitle(ByVal rng As Range, ByVal text As String)
    With rng.Cells(1, 1)
        .Value = text
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = ClrAccent()
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = ClrAccent()
        .Weight = xlThin
    End With
End Sub

Public Sub HeaderRow(ByVal rng As Range, ByVal headers As Variant)
    rng.Value = headers
    rng.Font.Bold = True
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = ClrAccent()
End Sub

Public Sub AddName(ByVal nm As String, ByVal rng As Range)
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & rng.Worksheet.Name & "'!" & rng.Address
End Sub
