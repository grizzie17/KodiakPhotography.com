<!--#include file="index_tools.asp"-->
<%



DIM	g_navtab_nFileCount
DIM	g_navtab_aFileList()
REDIM g_navtab_aFileList(5)
DIM	g_navtab_nIndex
DIM	g_navtab_sFile
DIM	g_navtab_sTitle


SUB navtabsGetList()

	DIM	sSaveSuffix
	DIM	nLen
	DIM	aFileSplit
	
	sSaveSuffix = g_sUseFileNameSuffix
	g_sUseFileNameSuffix = ".asp"
	buildFileListX g_navtab_nFileCount, g_navtab_aFileList, "./", FALSE
	g_sUseFileNameSuffix = sSaveSuffix
	
	
	g_navtab_sFile = ""
	g_navtab_nIndex = -1
	g_sPage = Request.ServerVariables("PATH_TRANSLATED")
	nLen = INSTRREV(g_sPage,"\",-1,vbTextCompare)
	IF 0 < nLen THEN g_sPage = MID(g_sPage,nLen+1)
	FOR nLen = 0 TO g_navtab_nFileCount-1
		aFileSplit = SPLIT( g_navtab_aFileList(nLen), vbTAB, -1, vbTextCompare )
		IF LCASE(g_sPage) = LCASE(aFileSplit(kFI_Name)) THEN
			g_navtab_sTitle = aFileSplit(kFI_Title)
			g_navtab_sFile = aFileSplit(kFI_Path)
			g_navtab_nIndex = nLen
			EXIT FOR
		END IF
	NEXT 'nLen
	IF 0 = LEN(g_navtab_sFile) THEN
		IF 0 < g_navtab_nFileCount THEN
			aFileSplit = SPLIT( g_navtab_aFileList(0), vbTAB, -1, vbTextCompare )
			g_sPage = aFileSplit(kFI_Name)
			g_navtab_sTitle = aFileSplit(kFI_Title)
			g_navtab_sFile = aFileSplit(kFI_Path)
			g_navtab_nIndex = 0
		ELSE
			g_sPage = "?????"
			g_navtab_sTitle = "??????"
			g_navtab_sFile = "?????"
			g_navtab_nIndex = 0
		END IF
	END IF

END SUB

navtabsGetList





%>