<% Response.Expires = 0
HTProgCap="報名表管理"
HTProgFunc="新增"
HTUploadPath="/public/"
HTProgCode="PA001"
HTProgPrefix="paAct" %>
<!--#Include file="../inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
	dim xNewIdentity
 apath=server.mappath(HTUploadPath) & "\"
' response.write apath & "<HR>"
  set xup = Server.CreateObject("UpDownExpress.FileUpload")
  xup.Open 

if xUpForm("submitTask") = "ADD" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		showErrBox()
	else
		doUpdateDB()
		showDoneBox()
	end if

else

	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if


function xUpForm(xvar)
on error resume next
	xStr = ""
	arrVal = xup.MultiVal(xvar)
	for i = 1 to Ubound(arrVal)
		xStr = xStr & arrVal(i) & ", "
'		Response.Write arrVal(i) & "<br>" & Chr(13)
	next 
	if xStr = "" then
		xStr = xup(xvar)
		xUpForm = xStr
	else
		xUpForm = left(xStr, len(xStr)-2)
	end if
end function
%>
<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- 新增時的表單預設值放在這裡
end sub

sub window_onLoad
	clientInitForm
end sub

    sub initRadio(xname,value)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    end sub

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()	
    

  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="ppPsnInfoFormDesign.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="ppPsnInfoQuery.asp" title="重設查詢條件">查詢</A>
		<%end if%>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=90% height=230 valign=top>    
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
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
        id=xUpForm("id")
        sql="select * from paInfoDesign where id=" & id 
        set rs=conn.Execute(sql)
        if rs.eof then
          sql = "INSERT INTO paInfoDesign values(" & xUpForm("id") & ",'" & xUpForm("check1_1")& "','" & xUpForm("check2_1")& "','" & xUpForm("check1_2")& "','" & xUpForm("check2_2")& "','" & xUpForm("check1_3")& "','" & xUpForm("check2_3")& "','" & xUpForm("check1_4")& "','" & xUpForm("check2_4")& "','" & xUpForm("check1_5")& "','" & xUpForm("check2_5")& "','" & xUpForm("check1_6")& "','" & xUpForm("check2_6")& "','" & xUpForm("check1_7")& "','" & xUpForm("check2_7")& "','" & xUpForm("check1_8")& "','" & xUpForm("check2_8")& "','" & xUpForm("check1_9")& "','" & xUpForm("check2_9")& "','" & xUpForm("check1_10")& "','" & xUpForm("check2_10")& "','" & xUpForm("check1_11")& "','" & xUpForm("check2_11")& "','" & xUpForm("check1_12")& "','" & xUpForm("check2_12")& "','" & xUpForm("check1_13")& "','" & xUpForm("check2_13")& "','" & xUpForm("check1_14")& "','" & xUpForm("check2_14")& "','" & xUpForm("check1_15")& "','" & xUpForm("check2_15")& "','" & xUpForm("check1_16")& "','" & xUpForm("check2_16")& "','" & xUpForm("check1_17")& "','" & xUpForm("check2_17")& "','" & xUpForm("check1_18")& "','" & xUpForm("check2_18")& "','" & xUpForm("check1_19")& "','" & xUpForm("check2_19")& "','" & xUpForm("check1_20")& "','" & xUpForm("check2_20")& "')"

          conn.Execute(sql)
        else
          sql="update paInfoDesign set check1_1='" & xUpForm("check1_1")& "',check2_1='" & xUpForm("check2_1")& "',check1_2='" & xUpForm("check1_2")& "',check2_2='" & xUpForm("check2_2")& "',check1_3='" & xUpForm("check1_3")& "',check2_3='" & xUpForm("check2_3")& "',check1_4='" & xUpForm("check1_4")& "',check2_4='" & xUpForm("check2_4")& "',check1_5='" & xUpForm("check1_5")& "',check2_5='" & xUpForm("check2_5")& "',check1_6='" & xUpForm("check1_6")& "',check2_6='" & xUpForm("check2_6")& "',check1_7='" & xUpForm("check1_7")& "',check2_7='" & xUpForm("check2_7")& "',check1_8='" & xUpForm("check1_8")& "',check2_8='" & xUpForm("check2_8")& "',check1_9='" & xUpForm("check1_9")& "',check2_9='" & xUpForm("check2_9")& "',check1_10='" & xUpForm("check1_10")& "',check2_10='" & xUpForm("check2_10")& "',check1_11='" & xUpForm("check1_11")& "',check2_11='" & xUpForm("check2_11")& "',check1_12='" & xUpForm("check1_12")& "',check2_12='" & xUpForm("check2_12")& "',check1_13='" & xUpForm("check1_13")& "',check2_13='" & xUpForm("check2_13")& "',check1_14='" & xUpForm("check1_14")& "',check2_14='" & xUpForm("check2_14")& "',check1_15='" & xUpForm("check1_15")& "',check2_15='" & xUpForm("check2_15")& "',check1_16='" & xUpForm("check1_16")& "',check2_16='" & xUpForm("check2_16")& "',check1_17='" & xUpForm("check1_17")& "',check2_17='" & xUpForm("check2_17")& "',check1_18='" & xUpForm("check1_18")& "',check2_18='" & xUpForm("check2_18")& "',check1_19='" & xUpForm("check1_19")& "',check2_19='" & xUpForm("check2_19")& "',check1_20='" & xUpForm("check1_20")& "',check2_20='" & xUpForm("check2_20")& "' where id=" &  xUpForm("id")
	  conn.Execute(sql)
	end if
	

end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("新增完成！")
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
