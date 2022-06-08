<%@ CodePage = 65001 %><%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/9 上午 10:55:22
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\web\subject\ct.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("kpi", "xItem", "ctNode", "mp", "ttxsl")
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
response.charset = "utf-8"
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1%>
<!--#include virtual = "/inc/checkURL.inc" -->
<!--#include virtual = "/subject/include/CheckPoint_BeforeLoadXML.asp" -->
<?xml version="1.0"  encoding="utf-8" ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
call CheckURL(Request.QueryString)
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	'---start of kpi use---20080911---vincent---
	if request.querystring("kpi") = "" then
		response.redirect "/Kpi/KpiTopicBrowse.aspx?xItem=" & request.querystring("xItem") & "&ctNode=" & request.querystring("ctNode") & "&mp=" & request.querystring("mp")
		response.end
	end if
	'---end of kpi use---
	mp = request("mp")
	if mp="" then 
		mp = session("mpTree")
	else
		session("mpTree") = mp
	end if
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	qStr = request.queryString
	if instr(qStr,"mp=")=0 then qStr = qStr & "&mp=" & session("mpTree")
	'若是文章沒有id則導回首頁  → 改成show單元建置中的視窗 回原頁
	aa= request("xItem")
	if aa = "" then 
	    response.write "<script>alert('單元建置中！');history.back();</script>"
		'response.redirect "/subject/mp.asp?mp=" &request("mp")
	end if

	memID = session("memID")
	gstyle = session("gstyle")
	
'this function in include file
call CheckPoint_BeforeLoadXML(request("mp"), request("ctNode"), request("xItem"))	
	
	
	xv = oxml.load(session("myXDURL") & "/wsxd/xdcp.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)	
	
	'response.write session("myXDURL") & "/wsxd/xdcp.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle & "<HR>"
	'response.end
  if oxml.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
  	
	xmyStyle= nullText(oxml.selectSingleNode("//hpMain/MpStyle"))
	if xmyStyle="" then	
		xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  	if xmyStyle="" then xmyStyle=session("myStyle")
	end if
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/cp.xsl"))
	fStyle = oxml.selectSingleNode("//xslData").text

	'response.write session("myXDURL") & "/wsxd/xdcp.asp?" & qStr & "<HR>"
	'response.end	  	
  if fStyle <> "" then
		set fxsl = server.createObject("microsoft.XMLDOM")
		fxsl.load(server.mappath("xslGip/" & xmyStyle & "/" & fStyle & ".xsl"))		
		'response.write (server.mappath("xslGip/" & xmyStyle & "/" & fStyle & ".xsl"))		
		set oxRoot = oxsl.selectSingleNode("xsl:stylesheet")
	
		on error resume next
		for each xs in fxsl.selectNodes("//xsl:template")
			set nx = xs.cloneNode(true)
			ckStr = "@match='" & nx.getAttribute("match") & "'"
			if nx.getAttribute("mode")<>"" then		ckStr = ckStr & " and @mode='" & nx.getAttribute("mode") & "'"
			set orgEx = oxRoot.selectSingleNode("//xsl:template[" & ckStr & "]")
			oxRoot.removechild orgEx
			oxRoot.appendChild nx
		next
		for each xs in fxsl.selectNodes("//msxsl:script")
			set nx = xs.cloneNode(true)
			oxRoot.appendChild nx
		next
	end if  

'	oxsl.load("http://chen2/chen/train/xml/xsltry/"&request("ttxsl"))
'	response.write oxml.xml
'	response.end
	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
    outString = replace(outString,"class=""lpphoto2""","class=""lpphoto2"" style='height:180px;overflow:hidden'")    
    outString = replace(outString,"2007 All Rights Reserved",year(date) & " All Rights Reserved")
	outString=replace(outString,"&amp;","&")
	outString=replace(outString,"&lt;","<")
	outString=replace(outString,"&gt;",">")			
	response.write outString
'	response.write oxml.transformNode(oxsl)
'	response.write oxml.xml & "<HR>"
'	response.write oxsl.xml & "<HR>"

Dim StrUrl
StrUrl=request.ServerVariables("HTTP_REFERER")
%>
<script type="text/javascript" src="/ViewCounter.aspx" ></script>
<script src="/js/jquery-ui-1.7.3.custom.min.js" type="text/javascript"></script>
<script>    var GB_ROOT_DIR = "/js/greybox/"; var referer_url = "<%=server.urlencode(StrUrl) %>";</script>
<script src="/js/greybox.js"></script>
<script src="/js/greybox/AJS.js"></script>
<script src="/js/greybox/gb_scripts.js"></script>
<script src="/TreasureHunt/treasurebox.js"></script>
