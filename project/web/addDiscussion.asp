<%
activeLog4U=true
onErrorPath="/"
progPath="D:\hyweb\GENSITE\project\web\addDiscussion.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("type2")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array()
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage
ChkNumeric chkNumericArray, chkNumericMessage
ChkDate chkDateArray, chkDateMessage
'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=request(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=request(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=request(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then response.redirect(onErrorPath)
	response.end
End Sub

'Log 放在網站根目錄下的 /Logs，檔名： YYYYMMDD_log4U.txt
Function Log4U(strLog)
	if (activeLog4U) then
		fldr=Server.mapPath("/") & "/Logs"
		filename=Year(Date()) & Right("0"&Month(Date()), 2) & Right("0"&Day(Date()),2)
		
		filename = filename & "_log4U.txt"
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'產生新的目錄
		If (Not fso.FolderExists(fldr)) Then
			Set f = fso.CreateFolder(fldr)
		Else
			ShowAbsolutePath = fso.GetAbsolutePathName(fldr)
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'開啟檔案
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile( fldr & "\" & filename , ForAppending, True, -1)
		f.Write strLog  & vbCrLf
	end if
End Function
'##### 結束：此段程式碼為自動產生，註解部份請勿刪除 #####
%><%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<html xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:hyweb="urn:gip-hyweb-com" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>農業知識入口網 －小知識串成的大力量－/</title>
</head>
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/client.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->

<%

'// purpose: decode a string
'// ex: ret = unicodeDecode(inputString, publicKey)
function unicodeDecode(byval inputString, byval publicKey)
    dim iPos
    dim returnValue
    dim currentChar

    returnValue = ""
    currentChar = ""

    if inputString <> "" then
        for iPos = 1 to len(inputString)
            currentChar = mid(inputString,iPos,1)
            'response.write "<br>" & iPos & ":" & currentChar & " --> " & uncodDecodeChar(currentChar, publicKey)
            returnValue = returnValue & uncodDecodeChar(currentChar, publicKey)
        next
    end if

    unicodeDecode = returnValue
end function

'// purpose: decode a char
'// ex: ret = unicodeDecodeChar(inputChar, publicKey)
function uncodDecodeChar(byval inputChar, byval publicKey)
    dim iPos
    dim returnValue
    dim currentByte

    returnValue = ""
    currentByte = ""

    if inputChar <> "" then
        for iPos = 1 to lenb(inputChar)
            currentByte = midb(inputChar,iPos,1)
            if iPos=1 then
                currentByte = chrb(ascb(currentByte)-publicKey)
            end if
            returnValue = returnValue & currentByte
        next
    end if
    
    uncodDecodeChar = returnValue
end function

	Function nullText(xNode)
  		On Error Resume Next
  		xstr = ""
  		xstr = xNode.text
  		nullText = xStr
	End Function
	
%>
<% 	if Session("memID")="" then%>
<script language="javascript">
	alert("請先登入會員");
	window.location.href=document.referrer;	
</script>
<% 	else %>

	<%
		if ((Request("CheckCode") <> Session("CheckCode")) and (Request("CheckCode")<>"true")) then
	%>
	<script language="javascript">
		alert("圖片驗證碼錯誤，請重新填寫");
		window.location.href=document.referrer;	
	</script>	  
	
<%
		else
		
			txtDisCussion = stripHTML(request("txtDisCussion"))
			txtDisCussion=Replace(txtDisCussion,vbCrLF, "<BR/>")
			'response.write txtDisCussion
			
			orderArticle = "1"
			IEditor = Session("memID")
			IDept = "0"
			showType = "1"
			siteId = "1"
			iBaseDSD = "44"
			iCTUnit = "2201"
			CtRootId = 0
			
			'2010/12/15 抓取知識拼圖的分數來計算 sam
			jigsawGrade = 1
			jigsawGradesql = " SELECT Rank0_1 FROM kpi_set_score WHERE (Rank0 = 'st_4') AND (Rank0_2 = 'st_419')  "
			set rsw = conn.Execute(jigsawGradesql)
			If Not rsw.EOF Then
				jigsawGrade = rsw("Rank0_1")
			End If
			insert1=0   '新增的超連結，沒有原始文章 

			title = request("title") 
			xItem=request("xItem")
			'parentIcuitem=request("xItem")

			sql1="SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.icuitem = KnowledgeJigsaw.gicuitem WHERE KnowledgeJigsaw.parentIcuitem = " & xItem & " AND CuDTGeneric.topCat = 'F'"
			set rs1=conn.Execute(sql1)
			parentIcuitem=rs1("icuitem")
			'response.write parentIcuitem
			'response.end()
			
			sql2 = "declare @newIDENTITY bigint"
			sql2 = sql2 & vbcrlf & "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[dEditDate],[iDept],[showType],[siteId],[xBody],[fCTUPublic]) "
			sql2 = sql2 & vbcrlf & "VALUES(" & iBaseDSD & ", " & iCTUnit & ", '" & title & "', '" & IEditor & "', GETDATE(), '" & IDept & "', '" & showType & "', '" & siteId & "','" & txtDisCussion & "','Y') "
			sql2 = sql2 & vbcrlf & "set @newIDENTITY = @@IDENTITY "
			sql2 = sql2 & vbcrlf & ""
			sql2 = sql2 & vbcrlf & "INSERT INTO CuDTx7 ([giCuItem]) VALUES(@newIDENTITY)"
			sql2 = sql2 & vbcrlf & ""
			sql2 = sql2 & vbcrlf & "INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[CtRootId],[CtUnitId],[parentIcuitem],[ArticleId],[Status],[orderArticle],[path]) "
			sql2 = sql2 & vbcrlf & "VALUES(@newIDENTITY, " & CtRootId & ", 1, " & parentIcuitem & ", " & insert1 & ", 'Y', " & orderArticle & ", '" & Url & "')"			
			'added by Joey, 增加留言增加KPI 分享度(互動)	
			'modified by Joey, http://gssjira.gss.com.tw/browse/COAKM-19, 將KPI的計算獨立到shareJigsaw欄位
			sql3=" declare @today datetime "
			sql3=sql3 & " declare @memberid varchar(50) "
			sql3=sql3 & " set @today = convert(varchar,getdate(), 111) "
			sql3=sql3 & " set @memberid='" & IEditor & "' "
			sql3=sql3 & " begin tran " & vbcrlf & " if (exists (select * from MemberGradeShare (updlock) where memberId=@memberid and convert(varchar,shareDate, 111)= convert(varchar,getdate(), 111))) "
			sql3=sql3 & " begin " & vbcrlf & " update MemberGradeShare set [shareJigsaw]=[shareJigsaw]+" & jigsawGrade & " where memberId=@memberid and convert(varchar,shareDate, 111)= convert(varchar,getdate(), 111) " & vbcrlf & "end "
			sql3=sql3 & " else" & vbcrlf & "begin" & vbcrlf
			sql3=sql3 & " insert into MemberGradeShare (memberID, shareDate,shareJigsaw) values(@memberid, getdate()," & jigsawGrade & ") end "
			
			sql2 =sql2 & sql3 & " commit"
			'response.write "sql=" & sql2
			conn.execute sql2 
		
%>
<script language="javascript">
	alert("留言成功");	
	window.location.href=document.referrer;	
	//history.back();
</script>
<%
		end if	
	end if
%>
