Sub gsx_rma_list()
    
    mTime = Time

    Worksheets("通取超過3日").UsedRange.ClearContents
    Worksheets("未關單").UsedRange.ClearContents
    Worksheets("未改待取").UsedRange.ClearContents

    If Worksheets("GSX").Range("F1").Value = "RMA單號" Then
    
    Else
        Worksheets("GSX").Select
        Worksheets("GSX").Columns("F:F").Select
        Selection.Insert Shift:=xlToRight
        Selection.Insert Shift:=xlToRight
        Selection.Insert Shift:=xlToRight
        Selection.Insert Shift:=xlToRight
        Selection.Insert Shift:=xlToRight

        Worksheets("GSX").Range("F1").Value = "RMA單號"
        Worksheets("GSX").Range("G1").Value = "RMA狀態"
        Worksheets("GSX").Range("H1").Value = "完修"
        Worksheets("GSX").Range("I1").Value = "取件"
        Worksheets("GSX").Range("J1").Value = "門市到件"

    End If
    
    
    GsxrowsCounts = Worksheets("GSX").Range("A1").CurrentRegion.Rows.Count
    GsxlineCounts = Worksheets("GSX").Range("A1").CurrentRegion.Columns.Count

    RmarowsCounts = Worksheets("RMA").Range("A1").CurrentRegion.Rows.Count
    RmalineCounts = Worksheets("RMA").Range("A1").CurrentRegion.Columns.Count

    Dim gsx_array, rma_array As Variant

    gsx_array = Worksheets("GSX").[a1].CurrentRegion
    rma_array = Worksheets("RMA").[a1].CurrentRegion

    For Lineindex = 1 To GsxlineCounts Step 1
        Worksheets("通取超過3日").Cells(1, Lineindex).Value = gsx_array(1, Lineindex)
        Worksheets("未關單").Cells(1, Lineindex).Value = gsx_array(1, Lineindex)
        Worksheets("未改待取").Cells(1, Lineindex).Value = gsx_array(1, Lineindex)
    Next Lineindex
    k = 2
    l = 2
    m = 2
     'MsgBox ("1st="&rma_array(2, 31))
    For i = 2 To GsxrowsCounts Step 1
        gsx_array(i, 6) = Left(gsx_array(i, 5), 10)
        Worksheets("GSX").Cells(i, 6) = Left(gsx_array(i, 5), 10)
        if len(gsx_array(i,6))= 10 and left(gsx_array(i,6),2) = "SA" then

	        'MsgBox ("gsx_rmaID="&gsx_array(i, 6)).Interior.ColorIndex = 37
	        For j = 2 To RmarowsCounts Step 1
	            If gsx_array(i, 6) = rma_array(j, 1) Then
	                gsx_array(i, 7) = rma_array(j, 20) 'rma狀態'
	                Worksheets("GSX").Cells(i, 7) = gsx_array(i, 7)

	                gsx_array(i, 8) = rma_array(j, 31)  '完修時間'
	                Worksheets("GSX").Cells(i, 8) = gsx_array(i, 8)

	                'msgbox(rma_array(2,31))
	                gsx_array(i, 9) = rma_array(j, 34)  '取件時間'
	                Worksheets("GSX").Cells(i, 9) = gsx_array(i, 9)

	                gsx_array(i, 10) = rma_array(j, 33) '到件時間'
	                Worksheets("GSX").Cells(i, 10) = gsx_array(i, 10)

	                '通知取件(門市到件日)超過3日，today 11/23,11/19以前(包括11/19) '
	                If gsx_array(i, 10) <> "" And gsx_array(i, 10) < DateAdd("d", -3, Date) Then
	                    For Lineindex = 1 To GsxlineCounts Step 1
	                        Worksheets("通取超過3日").Cells(k, Lineindex).Value = gsx_array(i, Lineindex)
	                    Next Lineindex
	                    k = k + 1
	                End If

	                '未關單為rma狀態 顧客領回'
	                If gsx_array(i, 7) = "顧客領回" Then
	                    For Lineindex = 1 To GsxlineCounts Step 1
	                        Worksheets("未關單").Cells(l, Lineindex).Value = gsx_array(i, Lineindex)
	                    Next Lineindex
	                    l = l + 1
	                End If

	                '未改待取件，Rma狀態為抵達門市、工程師完成、寄送到門市且gsx維修狀態不為待取件'
	                If (gsx_array(i, 7) = "抵達門市" Or gsx_array(i, 7) = "工程師完成" Or gsx_array(i, 7) = "寄送到門市") And gsx_array(i, 4) <> "待取件" Then
	                    For Lineindex = 1 To GsxlineCounts Step 1
	                        Worksheets("未改待取").Cells(m, Lineindex).Value = gsx_array(i, Lineindex)
	                    Next Lineindex
	                    m = m + 1
	                End If

	                Exit For
	            End If
	        Next j
	    else 
	    	Worksheets("GSX").Cells(i, 6).Interior.ColorIndex = 37
        
    Next i
    over3day = k - 2
    nClose = l - 2
    nChangestate = m - 2

    mTime = Time - mTime
    mSec = Second(mTime)

    MsgBox ("處理完成 " & "花費時間" & mSec & "秒" & vbCrLf & "通知取件超過3日 " & over3day & "筆" & vbCrLf & "未關單 " & nClose & "筆" & vbCrLf & "未改待取件 " & nChangestate & "筆")
    
End Sub

Sub clearsheetData()

    Worksheets("通取超過3日").UsedRange.ClearContents
    Worksheets("未關單").UsedRange.ClearContents
    Worksheets("未改待取").UsedRange.ClearContents
    Worksheets("GSX").UsedRange.ClearContents
    Worksheets("RMA").UsedRange.ClearContents

    MsgBox ("資料清除完成")
End Sub

