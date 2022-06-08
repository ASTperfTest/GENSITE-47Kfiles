<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:37
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdsitemap.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("mp")
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
%><?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<MenuTitle /> 
 <CtUnitName>網站導覽</CtUnitName> 
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<% 

Dim RSKey
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if
  		
  		myTreeNode = 0
                upParent = 0

  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
  	response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>"
	response.write "<mp>"&request("mp")&"</mp>"
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	xSitemapID=nullText(refModel.selectSingleNode("Sitemap"))

%>	
	<Sitemap myTreeNode="<%=myTreeNode%>">
<%

	SQLCom = "select * from CatTreeRoot Where CtRootID = "& xRootID
	set RS = conn.execute(SqlCom)


	FtypeName = RS("CtRootName")

   sql="select distinct datalevel from CatTreeNode where CtRootID = " & xRootID
   set rs0=conn.Execute(sql)
   do while not rs0.eof   
    if rs0(0)="1" then
        SqlCom = "SELECT a.CatName AS xCat,a.*,b.redirectURL, b.newWindow,b.iBaseDSD,b.CtUnitKind "_
                & " FROM CatTreeNode a,CtUnit b"_  
                & " WHERE  b.ctUnitID=*a.CtUnitID and a.inUse='Y' AND  a.DataLevel=" & rs0(0) & " AND  a.CtRootID = " & xRootID  _
                & " Order by  a.CatShowOrder"
   else
 	SqlCom = "SELECT b.CatName AS xCat, a.*, c.redirectURL, c.newWindow, c.iBaseDSD , c.CtUnitKind " _
		& " FROM CatTreeNode AS a JOIN CatTreeNode AS b ON b.CtNodeID=a.DataParent" _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataLevel=" & rs0(0) & " AND a.CtRootID = " & xRootID  _
		& " Order by b.CatShowOrder, a.CatShowOrder"
    end if		
	'response.write SqlCom & "<HR>"

	set RS = conn.execute(SqlCom)
	xCat = ""
	while not RS.eof
	  if rs0(0)="1" then
		xCat = RS("xCat")
%>
           <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
            <Caption><%=xCat%></Caption>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	end if
%>   
          <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>        
          </Menucat> 
          
        
<%		
          else
		xUrl = ""
		xNewWindow = ""
		if isNumeric(RS("iBaseDSD")) then
			xUrl = "List.asp?ctNode="&RS("CtNodeID") & "&CtUnit=" & RS("CtUnitID") & "&BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
		elseif RS("redirectURL")<> "" then
			xUrl = RS("redirectURL")
			if RS("newWindow")="Y"  then	xNewWindow = " target=""_nwMof"""
		end if
%>
                
<% if xUrl<>"" then %>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	end if
%>        
          <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
            <Caption><%=RS("CatName")%></Caption>  
            <xCat><%=RS("xCat")%></xCat>
            <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat>  
<% else %>
	 <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
            <Caption><%=RS("CatName")%></Caption>  
            <xCat><%=RS("xCat")%></xCat>
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat> 
<% end if %>
               
<%		
           end if
		RS.moveNext
	wend
	
   rs0.Movenext
   loop	
%>	

        </Sitemap> 
        
        	<Sitemap2 myTreeNode="<%=myTreeNode%>">
<%

	SQLCom = "select * from CatTreeRoot Where CtRootID = "& xSitemapID
	set RS = conn.execute(SqlCom)

        if not rs.eof then
	FtypeName = RS("CtRootName")

   sql="select distinct datalevel from CatTreeNode where CtRootID = " & xSitemapID
   set rs0=conn.Execute(sql)
   do while not rs0.eof   
    if rs0(0)="1" then
        SqlCom = "SELECT a.CatName AS xCat,a.*,b.redirectURL, b.newWindow,b.iBaseDSD,b.CtUnitKind "_
                & " FROM CatTreeNode a,CtUnit b"_  
                & " WHERE  b.ctUnitID=*a.CtUnitID and a.inUse='Y' AND  a.DataLevel=" & rs0(0) & " AND  a.CtRootID = " & xSitemapID  _
                & " Order by  a.CatShowOrder"
   else
 	SqlCom = "SELECT b.CatName AS xCat, a.*, c.redirectURL, c.newWindow, c.iBaseDSD , c.CtUnitKind " _
		& " FROM CatTreeNode AS a JOIN CatTreeNode AS b ON b.CtNodeID=a.DataParent" _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataLevel=" & rs0(0) & " AND a.CtRootID = " & xSitemapID  _
		& " Order by b.CatShowOrder, a.CatShowOrder"
    end if		
	'response.write SqlCom & "<HR>"

	set RS = conn.execute(SqlCom)
	xCat = ""
	while not RS.eof
	  if rs0(0)="1" then
		xCat = RS("xCat")
%>
           <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
            <Caption><%=xCat%></Caption>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	end if
%>   
          <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>        
          </Menucat> 
          
        
<%		
          else
		xUrl = ""
		xNewWindow = ""
		if isNumeric(RS("iBaseDSD")) then
			xUrl = "List.asp?ctNode="&RS("CtNodeID") & "&CtUnit=" & RS("CtUnitID") & "&BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
		elseif RS("redirectURL")<> "" then
			xUrl = RS("redirectURL")
			if RS("newWindow")="Y"  then	xNewWindow = " target=""_nwMof"""
		end if
%>
                
<% if xUrl<>"" then %>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	end if
%>        
          <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
            <Caption><%=RS("CatName")%></Caption>  
            <xCat><%=RS("xCat")%></xCat>
            <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat>  
<% else %>
	 <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
            <Caption><%=RS("CatName")%></Caption>  
            <xCat><%=RS("xCat")%></xCat>
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat> 
<% end if %>
               
<%		
           end if
		RS.moveNext
	wend
	
   rs0.Movenext
   loop	
   
   end if
%>	

        </Sitemap2> 


<%	
	for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
		processXDataSet
	next
 
  
function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function
%>

<!--#include file="x1Menus.inc" -->
</hpMain>
