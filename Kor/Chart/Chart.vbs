Sub start()
    Worksheets("Shop_data_summary").Range("A:A").Clear
    Worksheets("Shop_data_summary").Range("A2:AZ99").Clear
    Worksheets("Staff_data_summary").Range("C:C").Clear
    Worksheets("Staff_data_summary").Range("A2:AZ500").Clear
    Worksheets("testoo").Visible = True
    Worksheets("testoo").Range("A1:E999").Clear
    Worksheets("A_SalesData").Select
    ShopSummary
    Worksheets("A_SalesData").Select
    StaffSummary
    shopChart
    staffChart
    Worksheets("Chart").Select
    Worksheets("testoo").Visible = False
    Application.DisplayAlerts = True
    's = 3
    
    MsgBox ("Done!!")
    
End Sub


Sub DeleteDuplicates()
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = True
        Dim LastColumn As Integer
        LastColumn = Cells.Find(what:="*", after:=Range("C11"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("C10:C" & Cells(Rows.Count, 1).End(xlUp).Row)
            ' Use AdvanceFilter to filter unique values
            .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("staff_data_summary").Range("C1"), Unique:=True
            '.SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
            On Error Resume Next
            ActiveSheet.ShowAllData
            'Delete the blank rows
            'Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            Err.Clear
        End With
        'Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With
End Sub

Sub ShopSummary()
    
    Dim SaleDataRows As Long
    Dim ShopID, StaffID, Cat3id, Cat4id As String
    Dim Qty, CpuQty, iPadQty, watchQty, iPhoneQty, AirPodQty, StoreTotalMargin, Store3ppMargin, StoreMargin As Long
    Dim unit As Long
    unit = Worksheets("target").Range("I1").Value
    SaleDataRows = Worksheets("A_SalesData").Range("A11").CurrentRegion.Rows.Count - 1

    'tempStoreid = cells(11,1).value
    'Sheets("A_SalesData").Range(Cells(11, 1), Cells(SaleDataRows + 11, 1)).AdvancedFilter Action:=xlFilterCopy, Unique:=True

    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = True
        Dim LastColumn As Integer
        LastColumn = Cells.Find(what:="*", after:=Range("A11"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("A10:A" & Cells(Rows.Count, 1).End(xlUp).Row)

            .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("Shop_data_summary").Range("A1"), Unique:=True

            On Error Resume Next
            ActiveSheet.ShowAllData
            Err.Clear
        End With
        'Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With

    shopSummaryrow = Worksheets("shop_data_summary").Range("A1").CurrentRegion.Rows.Count

    Dim srcRange, fndRange As Range
    Dim fstAddress As String, i As Integer

    Worksheets("shop_data_summary").Activate
    Set srcRange = Worksheets("A_SalesData").Range("A11").CurrentRegion.Columns(1)



    For j = 2 To shopSummaryrow Step 1
        totalShop = 0
        Set fndRange = srcRange.Find(what:=Cells(j, 1))
        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            i = 1
            Do
                Cat3id = fndRange.Offset(, 4).Value
                Cat4id = fndRange.Offset(, 6).Value
                ShopQty = fndRange.Offset(, 8).Value
                ShopMargin = fndRange.Offset(, 9).Value / unit
                total3ppShop = fndRange.Offset(, 10) / unit
                Cells(j, 30).Value = Cells(j, 30).Value + total3ppShop
                Cells(j, 5).Value = Cells(j, 5).Value + ShopMargin

                Select Case Cat3id
                    Case Is = 3003
                        Cells(j, 7).Value = Cells(j, 7).Value + total3ppShop
                End Select

                Select Case Cat4id
                    Case Is = 4001 'CPU'
                        Cells(j, 13).Value = Cells(j, 13).Value + ShopQty
                    Case Is = 4002 'cpu'
                        Cells(j, 13).Value = Cells(j, 13).Value + ShopQty
                    Case Is = 4005 'ipad'
                        Cells(j, 18).Value = Cells(j, 18).Value + ShopQty
                    Case Is = 4006 'ipad mini'
                        Cells(j, 18).Value = Cells(j, 18).Value + ShopQty
                    Case Is = 4041 'ipad pro'
                        Cells(j, 18).Value = Cells(j, 18).Value + ShopQty
                    Case Is = 4004 'iphone'
                        Cells(j, 24).Value = Cells(j, 24).Value + ShopQty
                    Case Is = 4038 'watch'
                        Cells(j, 21).Value = Cells(j, 21).Value + ShopQty

                End Select

                Set fndRange = srcRange.FindNext(after:=fndRange)
                i = i + 1
            Loop Until fndRange.Address = fstAddress
            End If
            
            'cells(j,9).value = total3ppShop / totalshop
    Next j

    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)

    For k = 2 To shopSummaryrow Step 1
        Set fndRange = srcRange.Find(what:=Cells(k, 1))
        If Not fndRange Is Nothing Then
            Cells(k, 2).Value = fndRange.Offset(, 1).Value
            Cells(k, 3).Value = fndRange.Offset(, 2).Value
            Cells(k, 8).Value = fndRange.Offset(, 3).Value
            Cells(k, 11).Value = fndRange.Offset(, 4).Value
            Cells(k, 12).Value = fndRange.Offset(, 5).Value
            Cells(k, 14).Value = fndRange.Offset(, 6).Value
        End If
    Next k
   
   Set srcRange = Worksheets("B_ACPP+Data").Range("A1").CurrentRegion.Columns(1)

   For m = 2 To shopSummaryrow Step 1
        Set fndRange = srcRange.Find(what:=Cells(m, 1))
        Shoptotal = 0
        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            i = 1
            Do
                acppQty = 0
                cat5id = fndRange.Offset(, 4).Value
                cat6id = fndRange.Offset(, 6).Value
                acppQty = fndRange.Offset(, 8).Value
                'acppQty = fndRange.Offset(, 9).Value

                Select Case cat5id
                    Case Is = 5087
                        Cells(m, 27).Value = Cells(m, 27).Value + acppQty
                End Select
                Select Case cat6id
                    Case Is = 6524 'acpp cpu'
                        Cells(m, 16).Value = Cells(m, 16).Value + acppQty
                    Case Is = 6526  'acpp ipad'
                        Cells(m, 19).Value = Cells(m, 19).Value + acppQty
                    Case Is = 6528 'acpp watch'
                        Cells(m, 22).Value = Cells(m, 22).Value + acppQty
                    Case Is = 6523 'acpp ios iphone for kor'
                        Cells(m, 25).Value = Cells(m, 25).Value + acppQty
                    Case Is = 6525 'acpp iphone'
                        Cells(m, 25).Value = Cells(m, 25).Value + acppQty
                    Case Is = 6529  'acpp acc'
                        Cells(m, 28).Value = Cells(m, 28).Value + acppQty
                End Select
                Set fndRange = srcRange.FindNext(after:=fndRange)
                i = i + 1
            Loop Until fndRange.Address = fstAddress
        End If
        targetmg = Cells(m, 3).Value
        resultmg = Cells(m, 5).Value
        result3pp = Cells(m, 7).Value
        resultShopToral = Cells(m, 30).Value
        target3pp = Cells(m, 8).Value
        targetcpu = Cells(m, 14).Value
        resultcpu = Cells(m, 13).Value
        acppcpu = Cells(m, 16).Value
        ipadSale = Cells(m, 18).Value
        ipadacpp = Cells(m, 19).Value
        watchsale = Cells(m, 21).Value
        watchacpp = Cells(m, 22).Value
        iphonesale = Cells(m, 24).Value
        iphoneacpp = Cells(m, 25).Value
        airpodsale = Cells(m, 27).Value
        airpodacpp = Cells(m, 28).Value

        If targetmg = 0 Then
            MsgBox (Cells(m, 1).Value & " " & Cells(m, 2).Value & " Margin target is empty")
        Else
            Cells(m, 4).Value = targetmg * 0.8 'target80%
            Cells(m, 6).Value = resultmg / targetmg '毛利達成率'
        End If
        If resultmg = 0 Then
            MsgBox (Cells(m, 1).Value & " " & Cells(m, 2).Value & " 3pp margin result is empty")
        Else
            Cells(m, 9).Value = result3pp / resultShopToral '3pp搭_率''
            'Cells(m, 9).Value = cells(m,7) / Shoptotal
        End If
        If target3pp = 0 Then
            MsgBox (Cells(m, 1).Value & " " & Cells(m, 2).Value & " 3pp target is empty")
        Else
            Cells(m, 10).Value = Cells(m, 9).Value / target3pp
        End If
        If targetcpu = 0 Then
            MsgBox (Cells(m, 1).Value & " " & Cells(m, 2).Value & " cpu target is empty")
        Else
            Cells(m, 15).Value = resultcpu / targetcpu
        End If
        If resultcpu = 0 Then
            Cells(m, 17).Value = 0
        Else
            Cells(m, 17).Value = acppcpu / resultcpu
        End If
        If ipadSale = 0 Then
            Cells(m, 20).Value = 0
        Else
            Cells(m, 20).Value = ipadacpp / ipadSale
        End If
        If watchsale = 0 Then
            Cells(m, 23).Value = 0
        Else
            Cells(m, 23).Value = watchacpp / watchsale
        End If
        If iphonesale = 0 Then
            Cells(m, 26).Value = 0
        Else
            Cells(m, 26).Value = iphoneacpp / iphonesale
        End If
        If airpodsale = 0 Then
            Cells(m, 29).Value = 0
        Else
            Cells(m, 29).Value = airpodacpp / airpodsale
        End If
    Next m

    For o = 2 To shopSummaryrow Step 1
        totalmargin = totalmargin + Cells(o, 5).Value
        totaltarget = totaltarget + Cells(o, 3).Value
    Next o
        Cells(shopSummaryrow + 1, 2).Value = "Total"
        Cells(shopSummaryrow + 1, 4).Value = totaltarget * 0.8
        Cells(shopSummaryrow + 1, 3).Value = totaltarget
        Cells(shopSummaryrow + 1, 5).Value = totalmargin
        Cells(shopSummaryrow + 1, 6).Value = totalmargin / totaltarget

        Range("F:F,I:I,J:J,O:O,Q:Q,T:T,W:W,Z:Z,AC:AC").Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.0%"
        Range("C:E").NumberFormatLocal = "0.0"

    'MsgBox ("done")
End Sub

Sub StaffSummary()
    Dim SaleDataRows As Long
    Dim ShopID, StaffID, Cat3id, Cat4id As String
    Dim Qty, CpuQty, iPadQty, watchQty, iPhoneQty, AirPodQty, StoreTotalMargin, Store3ppMargin, StaffMargin As Long
    Dim unit As Long
    unit = Worksheets("target").Range("I1").Value
     SaleDataRows = Worksheets("A_SalesData").Range("A11").CurrentRegion.Rows.Count - 1

     With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = True
        Dim LastColumn As Integer
        LastColumn = Cells.Find(what:="*", after:=Range("C11"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("C10:C" & Cells(Rows.Count, 1).End(xlUp).Row)
            ' Use AdvanceFilter to filter unique values
            .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("Staff_data_summary").Range("C1"), Unique:=True
            '.SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
            On Error Resume Next
            ActiveSheet.ShowAllData
            'Delete the blank rows
            'Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            Err.Clear
        End With
        'Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With

    staffsummaryrows = Worksheets("Staff_data_summary").Range("C1").CurrentRegion.Rows.Count

    Dim srcRange, fndRange As Range
    Dim fstAddress As String, i As Integer
    'MsgBox (staffsummaryrows)
    Worksheets("Staff_data_summary").Activate
    Set srcRange = Worksheets("A_SalesData").Range("C11").CurrentRegion.Columns(3)
    Columns("A:A").Select

    Selection.NumberFormatLocal = "@"

    For j = 2 To staffsummaryrows Step 1
        Set fndRange = srcRange.Find(what:=Cells(j, 3))

        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            i = 1
            'MsgBox (fndRange.Offset(, -2).Value)
            'Cells(j, 1).formatlocal = "@"
            Cells(j, 1).Value = Format(fndRange.Offset(, -2).Text, "000")
            
           
            Cells(j, 2).Value = fndRange.Offset(, -1).Value
            Cells(j, 4).Value = fndRange.Offset(, 1).Value
            Do
                Cat3id = fndRange.Offset(, 2).Value
                Cat4id = fndRange.Offset(, 4).Value
                StaffQty = fndRange.Offset(, 6).Value
                StaffMargin = fndRange.Offset(, 7).Value
                StaffSaleSummary = fndRange.Offset(, 8).Value
                Cells(j, 5).Value = Cells(j, 5).Value + StaffSaleSummary / unit
                Cells(j, 6).Value = Cells(j, 6).Value + StaffMargin / unit

                Select Case Cat3id
                    Case Is = 3001
                        Cells(j, 8).Value = Cells(j, 8).Value + StaffQty
                End Select

                Set fndRange = srcRange.FindNext(after:=fndRange)
                i = i + 1
            Loop Until fndRange.Address = fstAddress

        End If
    Next j
    fndRange = Nothing
    Set srcRange = Worksheets("B_ACPP+Data").Range("C1").CurrentRegion.Columns(3)

    For k = 2 To staffsummaryrows Step 1
        Set fndRange = srcRange.Find(what:=Cells(k, 3))
        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            i = 1
            'Cells(k, 1).Value = fndRange.Offset(, -2).Value
            'Cells(k, 2).Value = fndRange.Offset(, -1).Value
            'Cells(k, 4).Value = fndRange.Offset(, 1).Value
            Do
                'acppQty = 0
                cat5id = fndRange.Offset(, 2).Value
                acppQty = fndRange.Offset(, 6).Value
                'msgbox(cat5id)
                Select Case cat5id
                    Case Is = 5090  'acpp+ '
                        Cells(k, 9).Value = Cells(k, 9).Value + acppQty
                End Select
                Set fndRange = srcRange.FindNext(after:=fndRange)
                i = i + 1
            Loop Until fndRange.Address = fstAddress
        End If

        staffTotalSale = Cells(k, 5).Value
        staffTotalMargin = Cells(k, 6).Value
        staffDev = Cells(k, 8).Value
        staffAcpp = Cells(k, 9).Value

        If staffTotalSale <> 0 Then
            Cells(k, 7).Value = staffTotalMargin / staffTotalSale
        End If
        If staffDev <> 0 Then
            Cells(k, 10).Value = staffAcpp / staffDev
        End If
    Next k

    Range("E:F").NumberFormatLocal = "0.0"
    Range("G:G,J:J").Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.0%"

    
    
    'fndRange = Nothing
    'MsgBox ("done")
End Sub

Sub shopChart()

    'Worksheets("Chart").ChartObjects.Delete
    k = False
    Application.DisplayAlerts = False
    For icount = 1 To Sheets.Count Step 1
        If Sheets(icount).Name = "Chart" Then
            k = True
        End If
    Next icount
    
    If k = True Then
            Worksheets("Chart").Delete
            Worksheets.Add
            ActiveSheet.Name = "Chart"
    Else
            Worksheets.Add
            ActiveSheet.Name = "Chart"
    End If
    Range("A1").Value = "Margin Target"
    Range("D1").Value = "Target"
    Range("E1").Select

    Selection.Value = Worksheets("target").Range("c1").Value
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.0%"

    Range("A2").Value = "Unit: " & Worksheets("target").Range("H1").Value
    Range("A15").Value = "3PP Attach Rate"
    Range("D15").Select
    Selection.Value = Worksheets("target").Range("F1").Value
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.0%"
    
    Range("M1").Value = "Apple (Unit: pcs)"
    Range("A32").Value = "Personal Performance"
    Range("C32").Value = "Unit: " & Worksheets("target").Range("H1").Value
    Range("A47").Value = "CPU (Unit: psc)"
    Range("A66").Value = "Apple Care + Attach % Perfornance"
    Range("A110").Value = "Personal performnce to sell Apple Care+ (Attach%)"
    Rows("13:13").Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.0%"

    Application.DisplayAlerts = False
    shopsummaryCount = Worksheets("shop_data_summary").Range("A1").CurrentRegion.Rows.Count - 1
    Dim gr As Range
    k = 0
    Worksheets("shop_data_summary").Select
    Set gr = Worksheets("Chart").Range("A3:C12")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
            '.Name = "TOTAL"
            .Chart.HasTitle = True
            .Chart.ChartTitle.Text = "TOTAL"
            .Chart.ChartTitle.Font.Size = 13
            '.Chart.SetSourceData Source:=Range("B1:E2")
            .Chart.SetSourceData Source:=Range("Shop_data_summary!$C$1:$E$1,Shop_data_summary!$C$" & shopsummaryCount + 1 & ":$E$" & shopsummaryCount + 1 & "")
            .Chart.FullSeriesCollection(1).ApplyDataLabels
            .Chart.HasAxis(xlValue, xlPrimary) = False
            .Chart.Axes(xlValue).HasMajorGridlines = False
            .Chart.HasLegend = False
    End With
    Worksheets("Chart").Range("A13:C13").Merge
    Worksheets("Chart").Range("A13").Value = Worksheets("Shop_data_summary").Range("F" & shopsummaryCount + 1 & "").Value
    'shop margin target'
    Dim mr As Range
    
    For i = 2 To shopsummaryCount Step 1
        Set gr = Worksheets("Chart").Range("D3:F12").Offset(, k)

        With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
            '.Name = Cells(2, i)
            .Chart.HasTitle = True
            .Chart.ChartTitle.Text = Cells(i, 2)
            .Chart.ChartTitle.Font.Size = 13
            '.Chart.SetSourceData Source:=Range("B1:E2")
            .Chart.SetSourceData Source:=Range("Shop_data_summary!$C$1:$E$1,Shop_data_summary!$C$" & i & ":$E$" & i & "")
            .Chart.FullSeriesCollection(1).ApplyDataLabels
            .Chart.HasAxis(xlValue, xlPrimary) = False
            .Chart.Axes(xlValue).HasMajorGridlines = False
            '.Chart.Axes(xlValue).MajorTickMark = false
            '=Shop_data_summary!$A$1,Shop_data_summary!$C$1:$E$1,Shop_data_summary!$A$3,Shop_data_summary!$C$3:$E$3
            .Chart.HasLegend = False
        End With

        k = k + 3

        Set gr = Nothing

    Next i
    Dim s, r As Integer
    
    s = 0
    '跨欄合_13欄''
    Worksheets("Chart").Select
    For j = 2 To shopsummaryCount Step 1
        'Set mr = Worksheets("Chart").Range("D13:F13").Offset(, s)
        'Worksheets("Chart").Range(Cells(13, 4 + s), Cells(13, 6 + s)).Merge across = True
        Worksheets("Chart").Range(Cells(13, 4 + s), Cells(13, 6 + s)).Select
        Selection.Merge across = True
        
        Worksheets("Chart").Cells(13, 4 + s).Value = Worksheets("Shop_data_summary").Range("F" & j & "").Value
        s = s + 3
        'Set mr = Nothing
    Next j
    Rows("13:13").Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.0%"
    Selection.HorizontalAlignment = xlCenter


    Worksheets("shop_data_summary").Select
    '3PP attach Rate'
    Set gr = Worksheets("Chart").Range("A16:L29")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$H$1:$J$" & shopsummaryCount & "")
        '.Chart.HasLegend = False
        .Chart.FullSeriesCollection(1).AxisGroup = 2
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(1).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(3).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.HasAxis(xlValue, xlPrimary) = False
        .Chart.HasAxis(xlValue, xlSecondary) = False
        .Chart.Axes(xlValue).HasMajorGridlines = False
       ' .Chart.Axes(xlValue).MajorGridlines = False
       .Chart.Legend.Position = xlLegendPositionBottom
       .Chart.SetElement (msoElementPrimaryValueAxisShow)
    .Chart.SetElement (msoElementSecondaryValueAxisShow)
        
    End With
    Set gr = Nothing
    'Apple 庫存與_滯''
    Worksheets("shop_data_summary").Select
    Set gr = Worksheets("Chart").Range("Ｍ16:U29")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$K$1:$L$" & shopsummaryCount & "")
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(1).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.HasAxis(xlValue, xlPrimary) = False
        .Chart.Axes(xlValue).HasMajorGridlines = False
        .Chart.Legend.Position = xlLegendPositionBottom
    End With
    Set gr = Nothing

    'CPU 銷_''
    Set gr = Worksheets("Chart").Range("A48:L64")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$M$1:$O$" & shopsummaryCount & "")
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(3).AxisGroup = 2
        
        .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(1).ApplyDataLabels
        .Chart.FullSeriesCollection(3).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.Legend.Position = xlLegendPositionBottom
        .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).MajorGridlines.Delete
        .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
       .Chart.Legend.Position = xlLegendPositionBottom
        
    End With

    'Apple Care + Attach % Perfornance
    Worksheets("shop_data_summary").Select
    'CPU ACPP+ attched'
    Set gr = Worksheets("Chart").Range("A67:I87")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$M$1:$M$" & shopsummaryCount & ",Shop_data_summary!$P$1:$Q$" & shopsummaryCount & "")
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "CPU"
        .Chart.ChartTitle.Font.Size = 13
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(3).AxisGroup = 2
        .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(1).ApplyDataLabels
        .Chart.FullSeriesCollection(3).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.Legend.Position = xlLegendPositionBottom
        .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).MajorGridlines.Delete
        .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
    End With

    'iPad ACPP+ Attched'
    Set gr = Worksheets("Chart").Range("J67:R87")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$R$1:$T$" & shopsummaryCount & "")
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "iPad"
        .Chart.ChartTitle.Font.Size = 13
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(3).AxisGroup = 2
        .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(1).ApplyDataLabels
        .Chart.FullSeriesCollection(3).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.Legend.Position = xlLegendPositionBottom
        .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).MajorGridlines.Delete
        .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
    End With

    'Watch ACPP+ Attched'
    Set gr = Worksheets("Chart").Range("S67:AA87")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$U$1:$W$" & shopsummaryCount & "")
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "Watch"
        .Chart.ChartTitle.Font.Size = 13
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(3).AxisGroup = 2
        .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(1).ApplyDataLabels
        .Chart.FullSeriesCollection(3).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.Legend.Position = xlLegendPositionBottom
        .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).MajorGridlines.Delete
        .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
    End With

    'iPhone ACPP+ Attched'
    Set gr = Worksheets("Chart").Range("A88:I107")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$X$1:$Z$" & shopsummaryCount & "")
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "iPhone"
        .Chart.ChartTitle.Font.Size = 13
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(3).AxisGroup = 2
        .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(1).ApplyDataLabels
        .Chart.FullSeriesCollection(3).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.Legend.Position = xlLegendPositionBottom
        .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).MajorGridlines.Delete
        .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
    End With

    'AirPodsACPP+ Attched'
    Set gr = Worksheets("Chart").Range("J88:R107")
    With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Chart.SetSourceData Source:=Range("Shop_data_summary!$B$1:$B$" & shopsummaryCount & ",Shop_data_summary!$AA$1:$AC$" & shopsummaryCount & "")
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "AirPods"
        .Chart.ChartTitle.Font.Size = 13
        .Chart.PlotBy = xlColumns
        .Chart.FullSeriesCollection(3).AxisGroup = 2
        .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
        .Chart.FullSeriesCollection(1).ApplyDataLabels
        .Chart.FullSeriesCollection(3).ApplyDataLabels
        .Chart.FullSeriesCollection(2).ApplyDataLabels
        .Chart.Legend.Position = xlLegendPositionBottom
        .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).TickLabelPosition = xlNone
        .Chart.Axes(xlValue).MajorGridlines.Delete
        .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
    End With


    'Worksheets("Char").Select
    
End Sub


Sub staffChart()
    
    'Worksheets("Chart").ChartObjects.Delete
    Worksheets("Staff_data_summary").Select

    SaleDataRows = Worksheets("Staff_data_summary").Range("A1").CurrentRegion.Rows.Count - 1
    'Action:=xlFilterCopy, Unique:=True
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = True
        Dim LastColumn As Integer
        LastColumn = Cells.Find(what:="*", after:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).Row)
            ' Use AdvanceFilter to filter unique values
            .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("testoo").Range("A1"), Unique:=True
            '.SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
            On Error Resume Next
            ActiveSheet.ShowAllData
            'Delete the blank rows
            'Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            Err.Clear
        End With
        'Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With

    Worksheets("testoo").Select
    testoorows = Worksheets("testoo").Range("A1").CurrentRegion.Rows.Count

    'Set srcRange = Worksheets("target").Range("A1").CurrentRegion.Columns(1)

    For i = 2 To testoorows Step 1
        Worksheets("testoo").Cells(i, 3).Value = "=VLOOKUP(A" & i & ",target!A:B,2,FALSE)"
    Next i
    
    Dim srcRange, fndRange As Range
    Dim fstAddress As String

    Worksheets("testoo").Activate
    Set srcRange = Worksheets("staff_data_summary").Range("A1").CurrentRegion.Columns(1)
    staffcharttest = Worksheets("testoo").Range("A1").CurrentRegion.Rows.Count

    '計算各個門市人員數'
    For j = 2 To staffcharttest Step 1
        Set fndRange = srcRange.Find(what:=Cells(j, 1))
        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            i = 1
            Do
                countNumber = countNumber + 1
                Worksheets("testoo").Cells(j, 2).Value = countNumber
                Set fndRange = srcRange.FindNext(after:=fndRange)
                i = i + 1
            Loop Until fndRange.Address = fstAddress
            countNumber = 0
        End If
    Next j

    'Personal Performance'
    j = 2
    countT = 0
    
    Dim gr As Range
    k = 0
    Worksheets("staff_data_summary").Select
    For i = 2 To staffcharttest Step 1
        ShopID = Worksheets("testoo").Cells(i, 1).Value
        'for j = j + countT to p step 1
        p = Worksheets("testoo").Cells(i, 2).Value
            Set gr = Worksheets("Chart").Range("A33:C45").Offset(, k)
            
            With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
                '.Name = Cells(1, i)
                '.Chart.SetSourceData Source:=Range("B1:E2")
                .Chart.HasTitle = True
                .Chart.ChartTitle.Text = Worksheets("testoo").Cells(i, 3)
                .Chart.ChartTitle.Font.Size = 12
                .Chart.SetSourceData Source:=Range("Staff_data_summary!$D$1:$G$1,Staff_data_summary!$D$" & j & ":$G" & j + p - 1 & "")
                '.Chart.HasLegend = False
                .Chart.FullSeriesCollection(3).AxisGroup = 2
                .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
                '.Chart.FullSeriesCollection(3).ApplyDataLabels
                .Chart.PlotBy = xlColumns
                '.Chart.HasAxis(xlValue, xlPrimary) = False
                '.Chart.HasAxis(xlValue, xlSecondary) = False
                .Chart.Axes(xlValue).TickLabelPosition = xlNone
                .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
                .Chart.Axes(xlValue).HasMajorGridlines = xlNone
                .Chart.FullSeriesCollection(1).ApplyDataLabels
                .Chart.FullSeriesCollection(2).ApplyDataLabels
                .Chart.FullSeriesCollection(3).ApplyDataLabels
                .Chart.Axes(xlValue).MajorTickMark = xlNone
                .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
                '.Chart.HasAxis(xlValue, xlPrimary) = False
                '.Chart.HasAxis(xlValue, xlSecondary) = False
                .Chart.Axes(xlValue).HasMajorGridlines = False
                .Chart.Axes(xlValue).MajorGridlines = False
               .Chart.Legend.Position = xlLegendPositionBottom
            End With

            'Personal performnce to Sell Apple Care +(Attach%)'
            Set gr = Worksheets("Chart").Range("A111:C122").Offset(, k)
            With Worksheets("Chart").ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
                '.Name = Cells(1, i)
                .Chart.SetSourceData Source:=Range("Staff_data_summary!$D$1,Staff_data_summary!$D$" & j & ":$D$" & j + p - 1 & ",Staff_data_summary!$H$1:$J$1,Staff_data_summary!$H$" & j & ":$J$" & j + p - 1 & "")
                .Chart.HasTitle = True
                .Chart.ChartTitle.Text = Worksheets("testoo").Cells(i, 3)
                .Chart.ChartTitle.Font.Size = 12
                .Chart.FullSeriesCollection(3).AxisGroup = 2
                .Chart.FullSeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
                
                .Chart.PlotBy = xlColumns

                .Chart.Axes(xlValue).TickLabelPosition = xlNone
                .Chart.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
                .Chart.Axes(xlValue).HasMajorGridlines = xlNone
                .Chart.FullSeriesCollection(1).ApplyDataLabels
                .Chart.FullSeriesCollection(2).ApplyDataLabels
                .Chart.FullSeriesCollection(3).ApplyDataLabels
                .Chart.Axes(xlValue).MajorTickMark = xlNone
                .Chart.Axes(xlValue, xlSecondary).MajorTickMark = xlNone

                .Chart.Axes(xlValue).HasMajorGridlines = False
                .Chart.Axes(xlValue).MajorGridlines = False
                .Chart.Legend.Position = xlLegendPositionBottom
            End With
            
            k = k + 3
            j = j + p
            Set gr = Nothing

        'next j
        'countT = worksheets("testoo").cells(i,2).value
        'p = p + countT
        '="Staff_data_summary!$D$1,Staff_data_summary!$D$" & j & ":$D$" & j + p - 1 & ",Staff_data_summary!$H$1:$J$1,Staff_data_summary!$H$" & j &":$J$" & j + p - 1&""
        '=Staff_data_summary!$D$1,Staff_data_summary!$D$14:$D$16,Staff_data_summary!$H$1:$J$1,Staff_data_summary!$H$14:$J$16
        '=Staff_data_summary!$D$1,Staff_data_summary!$D$14:$D$16,Staff_data_summary!$H$1:$J$1,Staff_data_summary!$H$14:$J$16
        '="Staff_data_summary!$H$1:$K$1,Staff_data_summary!$H$"& j &":$K$"& j + p - 1 &""
    Next i


End Sub

Sub ClearData()
    Worksheets("A_SalesData").Rows("11:99999").Delete
    Worksheets("B_ACPP+Data").Rows("2:99999").Delete
End Sub








