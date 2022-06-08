<%@ CodePage = 65001 %>
<%
'// purpose: to add epaper volume without user input.
'// modify date: 2006/01/06
'// ps: 1. 發行日為排程執行日.
'//     2. 標題為抓取 checkGIPconfig 中的參數.
'//     3. 起迄日為今天前7天至今天.
'//     4. 則數, 預設為5則.
'//
'// =========================================

Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="新增發行"
HTUploadPath="/public/"
HTProgCode="GW1M51"
HTProgPrefix="ePub" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<body>
<%
    dim defaultTitle
    dim defaultCount

    defaultTitle=getGIPconfigText("epaperDefaultTitle")
    defaultCount=getGIPconfigText("epaperDefaultCount")

	htx_pubDate = date
    htx_ePubID  = ""            '// empty, to identity.
    htx_title   = defaultTitle
    htx_dbDate  = date - 7
    htx_deDate  = date
    htx_maxNo   = defaultCount
	'checkDBValid()
	doUpdateDB htx_pubDate, htx_ePubID, htx_title, htx_dbDate, htx_deDate, htx_maxNo
%>
</body>
</html>


<% sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
  			window.history.back
	</script>
<%
end sub '---- showErrBox() ----


sub checkDBValid()	'===================== Server Side Validation Put HERE =================
'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "「客戶編號」重複!!請重新建立客戶編號!"
'		exit sub
'	end if
end sub '---- checkDBValid() ----

function getMaxVolume(byVal htx_title)
    dim rtnVal
    dim getSQL

	rtnVal = ""
	'getSQL = "Select * From epPub Where ctRootId = '"& session("epTreeID") &"' and "
	set RSvalidate = conn.Execute(getSQL)
	
	getMaxVolume = rtnVal
getMaxVolume
end function


sub doUpdateDB(byVal htx_pubDate, byVal htx_ePubID, byVal htx_title, byVal htx_dbDate, byVal htx_deDate, byVal htx_maxNo)
	sql = "INSERT INTO EpPub(ctRootId, "
	sqlValue = ") VALUES(" & session("epTreeID") & ","

	IF htx_pubDate <> "" Then
		sql = sql & "pubDate" & ","
		sqlValue = sqlValue & pkStr(htx_pubDate,",")
	END IF

	IF htx_ePubID <> "" Then
		sql = sql & "ePubID" & ","
		sqlValue = sqlValue & pkStr(htx_ePubID,",")
	END IF

	IF htx_title <> "" Then
		sql = sql & "title" & ","
		sqlValue = sqlValue & pkStr(htx_title,",")
	END IF

	IF htx_dbDate <> "" Then
		sql = sql & "dbDate" & ","
		sqlValue = sqlValue & pkStr(htx_dbDate,",")
	END IF

	IF htx_deDate <> "" Then
		sql = sql & "deDate" & ","
		sqlValue = sqlValue & pkStr(htx_deDate,",")
	END IF
	IF htx_maxNo <> "" Then
		sql = sql & "maxNo" & ","
		sqlValue = sqlValue & pkStr(htx_maxNo,",")
	END IF

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
	response.write sql
	'set RSx = conn.Execute(SQL)
	'xNewIdentity = RSx(0)
end sub 
'---- doUpdateDB() ----
%>
