'滿額抽獎名單'
'會員消費滿$2,023可獲得一次抽獎機會/ 消費滿$4,046有2次抽獎機會....... 消費滿$20,230有10次中獎機會，以此類推'
public  DateRange as Variant


sub lottery()
	worksheets("Start").select
	Dim basic As Long
	'取計算幾準E2的值'
	basic = Worksheets("Start").cells(2,5)

	worksheets("Temp").select
	RD_TempRows = Worksheets("Temp").Range("A1").CurrentRegion.Rows.Count
    RD_TempLins = Worksheets("Temp").Range("A1").CurrentRegion.Columns.Count
    Worksheets("Temp").Range(cells(1,1), Cells(RD_TempRows, RD_TempLins)).clear

	worksheets("090caa").select
    worksheets("090caa").range("A:A").copy 	destination:=Worksheets("Temp").Range("A1")
    
	worksheets("Temp").Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes

	DaterangeRows = Worksheets("Temp").Range("A1").CurrentRegion.Rows.Count
    DaterangeLins = Worksheets("Temp").Range("A1").CurrentRegion.Columns.Count

    '日期天數寫入陣列DateRange'
    worksheets("Temp").select
    redim DateRange(DaterangeRows)
    for i = 1 to DaterangeRows step 1
    	DateRange(i) = Worksheets("Temp").cells(i,1).Value
	next i

	Application.DisplayAlerts = False '關閉警告視窗

	For i = Worksheets.Count To  1 step -1
		for j = 1 to UBound(DateRange) step 1
			if worksheets(i).name = DateRange(j) then
				Worksheets(Worksheets.Count).Delete
			end if
		next j
	Next i

	Application.DisplayAlerts = True '恢復警告視窗


	for i = 2 to UBound(DateRange) step 1
       Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = DateRange(i)
       worksheets(DateRange(i)).range("A1").value = "消費日期"
       worksheets(DateRange(i)).range("B1").value = "消費門市"
       worksheets(DateRange(i)).range("C1").value = "會員後五碼"
       worksheets(DateRange(i)).range("D1").value = "電話後三碼"
       worksheets(DateRange(i)).range("E1").value = "消費金額"
       worksheets(DateRange(i)).range("F1").value = "會員代碼"
    next i

    worksheets("090caa").select
    RD_090caaRows = Worksheets("090caa").Range("A1").CurrentRegion.Rows.Count
    RD_090caaLins = Worksheets("090caa").Range("A1").CurrentRegion.Columns.Count
    worksheets("090caa").range("F1").value = "次數"
    
    for i = 2 to RD_090caaRows step 1
    	worksheets("090caa").cells(i,6).value = worksheets("090caa").cells(i,5).value \ basic
    next i
    RD_090caaLins = Worksheets("090caa").Range("A1").CurrentRegion.Columns.Count

    EPB090caa = Worksheets("090caa").Range(cells(1,1), Cells(RD_090caaRows, RD_090caaLins)).Value

    for i = 2 to UBound(EPB090caa) step 1
    	If EPB090caa(i,RD_090caaLins) > 0 Then
    		endrow = worksheets(EPB090caa(i,1)).cells(rows.count,1).end(xlup).row
    		for j = 1 to EPB090caa(i,RD_090caaLins) step 1
    			worksheets(EPB090caa(i,1)).cells(endrow+j,1).value = EPB090caa(i,1)	'消費日期'
    			worksheets(EPB090caa(i,1)).cells(endrow+j,2).value = EPB090caa(i,2)	'門市代碼'
    			worksheets(EPB090caa(i,1)).cells(endrow+j,5).value = EPB090caa(i,5)	'消費金額'
    			worksheets(EPB090caa(i,1)).cells(endrow+j,6).value = EPB090caa(i,3)	'會員代碼'
    			worksheets(EPB090caa(i,1)).cells(endrow+j,3).NumberFormatLocal = "@"
    			worksheets(EPB090caa(i,1)).cells(endrow+j,3).value = Format(Csng(right(EPB090caa(i,3),5)), "00000")
    			'worksheets(EPB090caa(i,1)).cells(endrow+j,3).value = Csng(right(EPB090caa(i,3),5)) '會員代碼取末五碼'
    			'if len(worksheets(EPB090caa(i,1)).cells(endrow+j,3).value)<3 then
    			''	worksheets(EPB090caa(i,1)).cells(endrow+j,3).Value = Format(, "00000")
    			'end if
    			worksheets(EPB090caa(i,1)).cells(endrow+j,4).NumberFormatLocal = "@"
    			if EPB090caa(i,4) = "" then 
    			else
    			worksheets(EPB090caa(i,1)).cells(endrow+j,4).value = Format(Csng(right(EPB090caa(i,4),3)), "000")
    			'worksheets(EPB090caa(i,1)).cells(endrow+j,4).value = Csng(right(EPB090caa(i,4),3)) '電話取末三碼'
    			end if
    		next j
    	End If
    next i
    msgbox "完成"
end sub


