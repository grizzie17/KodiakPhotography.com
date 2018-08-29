<%


'aList is a two dimensional array
'	index
'	0		use rank (identifies size)
'	1		importance rank (identifies color)
'	2		tag
SUB generateTagCloud( ByRef aList() )


	DIM	nHiUse
	DIM	nLoUse
	DIM	nHiImport
	DIM	nLoImport
	nHiUse = 0
	nLoUse = 999999
	nHiImport = 0
	nLoImport = 999999
	
	DIM	i
	FOR i = 0 to UBOUND(aList,0)
		IF nHiUse < aList(i,0) THEN nHiUse = aList(i,0)
		IF aList(i,0) < nLoUse THEN nLoUse = aList(i,0)
		IF nHiImport < aList(i,1) THEN nHiImport = aList(i,1)
		IF aList(i,1) < nLoImport THEN nLoImport = aList(i,1)
	NEXT

END SUB



%>