
Sub EDIINV()
    Worksheets("TempData").Range("A1:C999").Clear
    Worksheets("INV_CSV").Range("A:A").Clear
    Worksheets("APPLE_EDI_INV").Select

    Dim ExcelRow, INVRows, ShopCount, ModelCount As Integer
    Dim d, INVdata As Date
  
    INVRows = Worksheets("APPLE_EDI_INV").Range("A1").CurrentRegion.Rows.Count
    d = Worksheets("APPLE_EDI_INV").Range("I1").Value
    INVdata = worksheets("APPLE_EDI_INV").range("F1").value
    If Range("I1") = Range("J1") Then
        a = 0
    Else
        Worksheets("ShopID").Range("I1").Value = Format(1, "0000")
        Worksheets("APPLE_EDI_INV").Range("J1").Value = Range("I1").Value
    End If
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = True
        Dim LastColumn As Integer
        LastColumn = Cells.Find(what:="*", after:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).Row)
            ' Use AdvanceFilter to filter unique values
            .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("TempData").Range("A1"), Unique:=True
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

    Worksheets("TempData").Select

    Set srcRange = Worksheets("APPLE_EDI_INV").Range("A1").CurrentRegion.Columns(1)
    ShopCount = Worksheets("TempData").Range("A1").CurrentRegion.Rows.Count

    '計算各個門市人員數'
    For j = 2 To ShopCount Step 1
        Set fndRange = srcRange.Find(what:=Cells(j, 1))
        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            'i = 1
            Do
                countNumber = countNumber + 1
                Worksheets("TempData").Cells(j, 2).Value = countNumber
                Set fndRange = srcRange.FindNext(after:=fndRange)
                'i = i + 1
            Loop Until fndRange.Address = fstAddress
            countNumber = 0
        End If
    Next j

    'INVyy = DatePart("yy", d)
    'INVmm = DatePart("mm", d)
    'INVdd = DatePart("dd", d)
    INVyyyy = Format(d, "yyyy")
    INVyy = Format(d, "yy")
    INVmm = Format(d, "mm")
    INVdd = Format(d, "dd")
    reportdata_yyyy = Format(INVdata, "yyyy")
    reportdata_mm = Format(INVdata, "mm")
    reportdata_dd = Format(invdata, "dd")
    'msgbox(d)
    ms = Int(Rnd() * 9000 + 1000)
    SerialNumber = Format(Worksheets("ShopID").Range("I1").Value, "0000")
    
    HostID = Worksheets("ShopID").Range("E1").Value
    HQID = Worksheets("ShopID").Range("F1").Value
    'UNB+UNOA:3+總公司代碼:ZZ+060704780800T:01+年月日yymmdd:隨機四位數+流水號++INVRPT''
    'UNB+UNOA:3+X015715036XXX:ZZ+060704780800T:01+080408:0811+00000000000660++INVRPT''
    ExcelRow = 1
    Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "UNB+UNOA:3+" & HostID & ":ZZ+060704780800T:01+" & INVyyyy & INVmm & INVdd & ":" & ms & "+" & SerialNumber & "++INVRPT'"
    UNH_UNZ = 0
    
    Prod_run = 1
    lastProdCount = 0
    For i = 2 To ShopCount Step 1
        'UNH_UNZ = UNH_UNZ + 1

        BGM = Format(Worksheets("ShopID").Range("I2").Value, "0000")
        UNT = 0
        NAD_SHOPID = Worksheets("SHOPID").Cells(i, 2).Value
        ExcelRow = ExcelRow + 1
        Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "UNH+" & BGM & "+INVRPT:D:99A:UN'"
        'UNT = UNT + 1
        ExcelRow = ExcelRow + 1
        Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "BGM+35+" & BGM & "+9'"
        ExcelRow = ExcelRow + 1
        
        Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "DTM+137:" & INVyyyy & INVmm & INVdd & ":102'"
        ExcelRow = ExcelRow + 1
        Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "DTM+91:" & reportdata_yyyy & reportdata_mm & reportdata_dd & ":102'"
        ExcelRow = ExcelRow + 1
        Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "NAD+BY+" & HQID & "::91'"
        ExcelRow = ExcelRow + 1
        Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "NAD+GY+" & NAD_SHOPID & "::90'"
        ProdCount = Worksheets("TempData").Cells(i, 2).Value

        For j = 1 To ProdCount Step 1
            LIN_VP = Worksheets("APPLE_EDI_INV").Cells(j + 1 + lastProdCount, 2).Value
            QTY = Worksheets("APPLE_EDI_INV").Cells(j + 1 + lastProdCount, 3).Value
            ExcelRow = ExcelRow + 1
            Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "LIN+" & j & "++" & LIN_VP & ":VP'"
            ExcelRow = ExcelRow + 1
            Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "INV+2'"
            ExcelRow = ExcelRow + 1
            Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "QTY+17:" & QTY & "'"

        Next j
        Worksheets("ShopID").Range("I2").Value = Worksheets("ShopID").Range("I2").Value + 1
        ExcelRow = ExcelRow + 1
        UNT = ProdCount * 3 + 7
        Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "UNT+" & UNT & "+" & BGM & "'"
        lastProdCount = lastProdCount + ProdCount
        If BGM = 9999 Then
            Worksheets("ShopID").Range("I2").Value = 1
        End If
    Next i
    ExcelRow = ExcelRow + 1
    Worksheets("INV_CSV").Cells(ExcelRow, 1).Value = "UNZ+" & ShopCount - 1 & "+" & SerialNumber & "'"
    Worksheets("INV_CSV").Select
    
    MsgBox ("Please saveas in CSV ")
End Sub





