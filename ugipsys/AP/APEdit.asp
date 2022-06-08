<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="AP"
HTProgCode="HT003"
HTProgPrefix="AP" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim RSreg
Dim formFunction
taskLable="編輯" & HTProgCap

if request("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
	end if

elseif request("submitTask") = "DELETE" then
	SQL = "DELETE FROM Ap WHERE apcode=" & pkStr(request.queryString("APcode"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if


sub EditInBothCase
	showHTMLHead()
	sqlCom = "SELECT * FROM Ap WHERE apcode=" & pkStr(request.queryString("APcode"),"")
  Set RSreg = Conn.execute(sqlcom)

	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("pfx_"&fldName) <> "" then
			xValue = request("pfx_"&fldName)
		elseif request("tfx_"&fldName) <> "" then
			xValue = request("tfx_"&fldName)
		elseif request("dfx_"&fldName) <> "" then
			xValue = request("dfx_"&fldName)
		elseif request("sfx_"&fldName) <> "" then
			xValue = request("sfx_"&fldName)
		elseif request("nfx_"&fldName) <> "" then
			xValue = request("nfx_"&fldName)
		elseif request("bfx_"&fldName) <> "" then
			xValue = request("bfx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function 
%>

<%Sub initForm() %>
    <script language=vbs>
      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

     sub clientInitForm
	reg.pfx_apcode.value = "<%=qqRS("apcode")%>"
	reg.tfx_apnameE.value = "<%=qqRS("apnameE")%>"
	reg.tfx_apnameC.value = "<%=qqRS("apnameC")%>"
	reg.sfx_apcat.value = "<%=qqRS("apcat")%>"
	reg.tfx_appath.value = "<%=qqRS("appath")%>"
	reg.tfx_aporder.value = "<%=qqRS("aporder")%>"
	reg.nfx_apMask.value = "<%=qqRS("apmask")%>"
	reg.tfx_spare64.value = "<%=qqRS("spare64")%>"
	reg.tfx_spare128.value = "<%=qqRS("spare128")%>"
	reg.sfx_xsNewWindow.value = "<%=qqRS("xsNewWindow")%>"
	reg.sfx_xsSubmit.value = "<%=qqRS("xsSubmit")%>"
	initBitsArray <%=qqRS("apmask")%>

    end sub

    sub window_onLoad
         clientInitForm
    end sub
 </script>	
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
    <script language="vbscript">
      sub formModSubmit()
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>

<!--#include file="APForm.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;"> 
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
    <title>編修表單</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【<%=HTProgCap%>編輯】</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">查詢</a>&nbsp;
	       <%end if%>	    
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="<%=HTProgPrefix%>Add.asp">新增</a>&nbsp;
	       <%end if%>
	       <a href="Javascript:window.history.back();">回前頁</a>
		</td>
	
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>        
<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
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
<%End sub '--- showHTMLTail() ------%>


<%Sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
'  			window.history.back
	</script>
<%
    End sub '---- showErrBox() ----

Sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = '"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "「統一編號」重複!!請重新輸入客戶統一編號!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()
	sql = "UPDATE Ap SET "
	sqlWhere = ""
	for each x in request.form
	  if mid(x,2,3) = "fx_" then
		xfldName = mid(x,5)
		select case left(x,1)
		  case "p"
			if sqlWhere="" then 
				sqlWhere = " WHERE " & mid(x,5) & "=" & pkStr(request(x),"")
			else
				sqlWhere = sqlWhere & " AND " & mid(x,5) & "=" & pkStr(request(x),"")
			end if
		  case "d"
			sql = sql & " " & mid(x,5) & "=" & pkStr(request(x),",")
		  case "n"
			sql = sql & " " & mid(x,5) & "=" & drn(x)
		  case else
			sql = sql & " " & mid(x,5) & "=" & pkStr(request(x),",")
		end select
	  end if
	next 
	sql = left(sql,len(sql)-1) & sqlWHERE
'response.write sql
'response.end	
	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>編修表單</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=HTprogPrefix%>List.asp?page_no=<%=Session("QueryPage_No")%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
