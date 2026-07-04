'DOA率'
public  SalesTable,YM,DOAYM,FinalTable as Variant

sub Doa()

	worksheets("銷售資料").select
    worksheets("銷售資料").range("A:F").copy 	destination:=Worksheets("Temp").Range("A1")
    worksheets("Temp").select
	worksheets("Temp").Range("A:F").RemoveDuplicates Columns:=5, Header:=xlYes
	'msgbox worksheets("DOA率").cells(1,7).value
	finalrow = cells(rows.count,1).end(xlup).row
	worksheets("Temp").range(cells(2,1),cells(finalrow,6)).copy 	destination:=Worksheets("DOA率").Range("A3")
	'TempSalesRows = Worksheets("Temp").Range("A1").CurrentRegion.Rows.Count
    'TempSalesLins = Worksheets("Temp").Range("A1").CurrentRegion.Columns.Count
    'msgbox finalrow

    '銷售資料寫入陣列'
	SalesRows = Worksheets("銷售資料").Range("A1").CurrentRegion.Rows.Count
    SalesLins = Worksheets("銷售資料").Range("A1").CurrentRegion.Columns.Count
    worksheets("銷售資料").select
    SalesTable = Worksheets("銷售資料").Range(cells(1,1), Cells(SalesRows, SalesLins)).Value
    '銷售資料的年月名稱寫入'
    countM = SalesLins - 6
    ReDim YM(countM)
    for i = 1 to countM step 1
    	YM(i) = mid(Worksheets("銷售資料").cells(1,6+i).value,8,6)
    next i

    DoaRows = Worksheets("DOA退貨資料").Range("A1").CurrentRegion.Rows.Count
    DoaLins = Worksheets("DOA退貨資料").Range("A1").CurrentRegion.Columns.Count
    worksheets("DOA退貨資料").select
    DoaTable = Worksheets("DOA退貨資料").Range(cells(1,1), Cells(DoaRows, DoaLins)).Value
    'DOA資料的年月名稱寫入'
    countDoaM = SalesLins - 6
    ReDim DOAYM(countDoaM)
    for i = 1 to countDoaM step 1
    	DOAYM(i) = mid(Worksheets("DOA退貨資料").cells(1,2+i).value,14,6)
    next i

    FinalRows = Worksheets("DOA率").Range("A1").CurrentRegion.Rows.Count
    FinalLins = Worksheets("DOA率").Range("A2").CurrentRegion.Columns.Count

    'DOA率的年月欄位名稱寫入'
    redim x(12)
    i = 1
    for k = 7 to 40 step 3
    	x(i) = worksheets("DOA率").cells(1,k).value
    	i= i+ 1
    next 

    '比對年月與銷售資料寫入final表'
    for m = 1 to countM step 1
    	'salesym = YM(m)
    	for n = 1 to 12 step 1
    		if YM(m) = x(n) then
    			salesyyyymm = 6 + m
    			DOApercentYM = (n-1)*3 + 7
			    for i = 3 to FinalRows step 1
			    	for j = 1 to UBound(SalesTable) step 1
			    		if worksheets("DOA率").cells(i,5).value = SalesTable(j,5) then
			    			worksheets("DOA率").cells(i,DOApercentYM).value = SalesTable(j,salesyyyymm) + worksheets("DOA率").cells(i,DOApercentYM).value
			    		end if
			    	next j
			    next i
			end if	
		next n
	next m

	'比對年月與DOA資料寫入final表'
    for m = 1 to countDoaM step 1
    	'salesym = YM(m)
    	for n = 1 to 12 step 1
    		if DOAYM(m) = x(n) then
    			doayyyymm = 2 + m
    			DOApercentYM = (n-1)*3 + 8
			    for i = 3 to FinalRows step 1
			    	for j = 1 to UBound(DoaTable) step 1
			    		if worksheets("DOA率").cells(i,5).value = DoaTable(j,1) then
			    			worksheets("DOA率").cells(i,DOApercentYM).value = DoaTable(j,doayyyymm) + worksheets("DOA率").cells(i,DOApercentYM).value
			    		end if
			    	next j
			    next i
			end if	
		next n
	next m
	for n = 1 to 12 step 1
		finalDoaYM = (n-1)*3 + 9
		for i = 3 to FinalRows step 1
			if worksheets("DOA率").cells(i,finalDoaYM-2).value = 0 then
				worksheets("DOA率").cells(i,finalDoaYM).value = "-"
			else
				worksheets("DOA率").cells(i,finalDoaYM).value = worksheets("DOA率").cells(i,finalDoaYM-1).value / worksheets("DOA率").cells(i,finalDoaYM-2).value
			end if
		next i
	next n
	'每季銷量欄位位置'
	Q1 = 43
	Q2 = 46
	Q3 = 49
	Q4 = 52
	allday = 55

	Worksheets("DOA率").select
	FinalTable = Worksheets("DOA率").Range(cells(3,1), Cells(FinalRows, FinalLins)).Value
	
	for i = 3 to FinalRows step 1

		worksheets("DOA率").cells(i,Q1).value = FinalTable(i,7)+FinalTable(i,10)+FinalTable(i,13)
		worksheets("DOA率").cells(i,Q2).value = FinalTable(i,16)+FinalTable(i,19)+FinalTable(i,22)
		worksheets("DOA率").cells(i,Q3).value = FinalTable(i,25)+FinalTable(i,28)+FinalTable(i,31)
		worksheets("DOA率").cells(i,Q4).value = FinalTable(i,34)+FinalTable(i,37)+FinalTable(i,40)
		worksheets("DOA率").cells(i,allday).value = worksheets("DOA率").cells(i,Q1).value+worksheets("DOA率").cells(i,Q2).value+worksheets("DOA率").cells(i,Q3).value+worksheets("DOA率").cells(i,Q4).value

		worksheets("DOA率").cells(i,Q1+1).value = FinalTable(i,8)+FinalTable(i,11)+FinalTable(i,14)
		worksheets("DOA率").cells(i,Q2+1).value = FinalTable(i,17)+FinalTable(i,20)+FinalTable(i,23)
		worksheets("DOA率").cells(i,Q3+1).value = FinalTable(i,26)+FinalTable(i,29)+FinalTable(i,32)
		worksheets("DOA率").cells(i,Q4+1).value = FinalTable(i,35)+FinalTable(i,38)+FinalTable(i,41)
		worksheets("DOA率").cells(i,allday+1).value = worksheets("DOA率").cells(i,Q1+1).value+worksheets("DOA率").cells(i,Q2+1).value+worksheets("DOA率").cells(i,Q3+1).value+worksheets("DOA率").cells(i,Q4+1).value

		if FinalTable(i,Q1) = 0 then
			FinalTable(i,Q1+2) = "-"
		else
			FinalTable(i,Q1+2) = FinalTable(i,Q1+1) / FinalTable(i,Q1)
		end if
		if FinalTable(i,Q2) = 0 then
			FinalTable(i,Q2+2) = "-"
		else
			FinalTable(i,Q2+2) = FinalTable(i,Q2+1) / FinalTable(i,Q2)
		end if
		if FinalTable(i,Q3) = 0 then
			FinalTable(i,Q3+2) = "-"
		else
			FinalTable(i,Q3+2) = FinalTable(i,Q3+1) / FinalTable(i,Q3)
		end if
		if FinalTable(i,Q4) = 0 then
			FinalTable(i,Q4+2) = "-"
		else
			FinalTable(i,Q4+2) = FinalTable(i,Q4+1) / FinalTable(i,Q4)
		end if
		if FinalTable(i,allday) = 0 then
			FinalTable(i,allday+2) = "-"
		else
			FinalTable(i,allday+2) = FinalTable(i,allday+1) / FinalTable(i,allday)
		end if

	next i
    'msgbox x(3)
end sub