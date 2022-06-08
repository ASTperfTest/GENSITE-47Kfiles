<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="AP"
HTProgCode="HT003"
HTProgPrefix="AP" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
if request("submitTask") = "ADD" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		showErrBox()
	else
		doUpdateDB()
		showDoneBox()
	end if

else
	ClientID=0
	taskLable="新增" & HTProgCap

	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- 新增時的表單預設值放在這裡
	reg.nfx_APMask.value = "31"
	initBitsArray 31
end sub

sub window_onLoad
	clientInitForm
end sub
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()	
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="APForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【<%=HTProgCap%>新增】</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">查詢</a>	    
	       <%end if%>
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
	<!--#include virtual = "/inc/Footer.inc"-->
      </td>                                         
  </tr> 
</table> 
</body>
</html>
<% end sub '--- showHTMLTail() ------%>


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

sub doUpdateDB()
	sql = "INSERT INTO Ap("
	sqlValue = ") VALUES("
	for each x in request.form
	 if request(x) <> "" then
	  if mid(x,2,3) = "fx_" then
		select case left(x,1)
		  case "p"
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & pkStr(request(x),",")
		  case "d"
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & pkStr(request(x),",")
		  case "n"
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & request(x) & ","
		  case else
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & pkStr(request(x),",")
		end select
	  end if
	 end if
	next 
	if request("tfx_appath") = "" then
		sql = sql & "appath,specPath,"
		sqlValue = sqlValue & "'/HTSDSystem/ucVersion.asp?apCode=" & request("pfx_apcode") & "'," _
			& "'/HTSDSystem/ucVersion.asp?apCode=" & request("pfx_APcode") & "'," 
	else
		sql = sql & "specPath,"
		sqlValue = sqlValue & pkStr(request("tfx_appath"),",")
	end if
	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

'	response.write sql & "<HR>"
	conn.Execute(SQL)  

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/High-Level.xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if


'  response.write oxml.selectSingleNode("//frame[@name='topFrame']/@src").text

	oxml.selectSingleNode("UseCase/Name").text = request("tfx_apnameC")
	oxml.selectSingleNode("UseCase/Code").text = request("pfx_apcode")
	oxml.selectSingleNode("UseCase/APCat").text = request("sfx_apcat")
	oxml.selectSingleNode("UseCase/Version/Date").text = date()
	oxml.selectSingleNode("UseCase/Version/Author").text = session("userID")
'  response.write oxml.documentElement.xml
	oxml.save(Server.MapPath("/HTSDSystem/UseCase/" & request("pfx_apcode") & ".xml"))

'  response.end


end sub '---- doUpdateDB() ----

%>

<% sub showDoneBox() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("新增完成！")
	document.location.href = "<%=HTprogPrefix%>Query.asp"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
