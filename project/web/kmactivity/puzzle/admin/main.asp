<%
If session("login")<>"ok" Then
	session("errnumber")=1
  	session("msg")="請登入帳密！"
	Response.Redirect "login.asp"
End If

Sub Message()
  If session("errnumber")=0 then
    Response.Write "<center>"&session("msg")&"</center>"
  Else
    Response.Write "<script Language='JavaScript'>alert('"&session("msg")&"')</script>"
  End If
  session("msg")=""
  session("errnumber")=0
End Sub

Function SubmitJ (SubmitType)
  if SubmitType="delete" then Response.Write "if(confirm('您是否確定要送出？')){" & Chr(13) & Chr(10)
  Response.Write "document.form.action.value='"&SubmitType&"';" & Chr(13) & Chr(10)
  Response.Write "document.form.submit();" & Chr(13) & Chr(10)
  if SubmitType="delete" then Response.Write "}" & Chr(13) & Chr(10)
End Function

Function GetSecureVal(param)
	If IsEmpty(param) Or param = "" Then
		GetSecureVal = param
		Exit Function
	End If

	If IsNumeric(param) Then
		GetSecureVal = CLng(param)
	Else
		GetSecureVal = Replace(CStr(param), "'", "''")
	End If
End Function
	
set conn = server.createobject("adodb.connection")
Conn.ConnectionString = application("ConnStrPuzzle")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open

If request("action")="add" Then
  SQL="PICDATA"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("Pic_No")=request.form("Pic_No_New")
  RS("Pic_Name")=request.form("Pic_Name_New")
  RS("Pic_Link")=GetSecureVal(request.form("Pic_Link_New"))
  RS("Pic_Open")=GetSecureVal(request.form("Pic_Open_New"))
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料新增成功 ！"
End If

If request("action")="update" Then	  
  SQL="Select * From PICDATA Where Ser_No='"&request.form("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Pic_No")=request.form("Pic_No"&"_"&Request("Ser_No"))
  RS("Pic_Name")=request.form("Pic_Name"&"_"&Request("Ser_No"))
  RS("Pic_Link")=GetSecureVal(request.form("Pic_Link"&"_"&Request("Ser_No")))
  RS("Pic_Open")=GetSecureVal(request.form("Pic_Open"&"_"&Request("Ser_No")))
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
End If

If request("action")="delete" Then
  SQL="Delete From PICDATA Where Ser_No='"&GetSecureVal(request.form("ser_no"))&"' "       
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="資料刪除成功 ！"
End If
	
SQL = "Select * From PICDATA Order by Ser_No"		
set rs = conn.execute (SQL)

%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=big5">
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
  <!--<link REL="stylesheet" type="text/css" HREF="../include/dms.css">-->
  <link href="main.css" rel="stylesheet" type="text/css" />
  <link href="table.css" rel="stylesheet" type="text/css" />
  <title>圖庫管理</title>
</head>
<body class="body">
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input name='ser_no' type='hidden' value=''>
      	<table width="100%" border=1 cellspacing="0" cellpadding="2" class="table_v">
        	<tr>
            	<th noWrap align="center">&nbsp;</th>
            	<th noWrap align="center">目錄名</th>
                <th noWrap align="center">圖檔主檔名</th>
                <th noWrap align="center">連結網址</th>
                <th noWrap align="center">圖片開放</th>
                <th noWrap align="center">&nbsp;</th>
            </tr>
            <tr>
            	<td>&nbsp;</td>
            	<td><input class=font9 size=20 name="Pic_No_New" maxlength="50" ></td>
                <td><input class=font9 size=20 name="Pic_Name_New" maxlength="50" ></td>
                <td><input class=font9 size=50 name="Pic_Link_New" maxlength="100" ></td>
                <td><input name="Pic_Open_New" type="radio" value="Y" checked>開放
                	<input name="Pic_Open_New" type="radio" value="N" >關閉</td>
                <td><input type="button" value=" 新 增 " name="add" class="delbutton" style="cursor:hand" OnClick='Add_OnClick();'></td> 
            </tr>
        	<tr>
            	<th noWrap align="center">圖示</th>
                <th noWrap align="center">目錄名</th>
                <th noWrap align="center">圖檔主檔名</th>
                <th noWrap align="center">連結網址</th>
                <th noWrap align="center">圖片開放</th>
                <th noWrap align="center">&nbsp;</th>
                <!--<th noWrap align="center">&nbsp;</th>-->
            </tr>
        <%
		  While Not RS.Eof%>
        	<tr>            	
                <td><img border="0" src="../puzzlePics/<%=RS("Pic_No")%>/<%=RS("Pic_Name")%>.jpg"></td>
                <td><input class=font9 size=20 name="Pic_No_<%=RS("Ser_No")%>" maxlength="50" value="<%=RS("Pic_No")%>"></td>
                <td><input class=font9 size=20 name="Pic_Name_<%=RS("Ser_No")%>" maxlength="50" value="<%=RS("Pic_Name")%>"></td>
                <td><input class=font9 size=50 name="Pic_Link_<%=RS("Ser_No")%>" maxlength="100" value="<%=RS("Pic_Link")%>"></td>
                <td><input name="Pic_Open_<%=RS("Ser_No")%>" type="radio" value="Y" <%If RS("Pic_Open")="Y" Then Response.write "checked" %>>開放
                	<input name="Pic_Open_<%=RS("Ser_No")%>" type="radio" value="N" <%If RS("Pic_Open")="N" Then Response.write "checked" %>>關閉</td>
                <td><input type="button" value=" 儲 存 " name="update" class="delbutton" style="cursor:hand" OnClick='Update_OnClick(<%=RS("Ser_No")%>);'></td>
                <!--<td><input type="button" value=" 刪 除 " name="del" class="addbutton" style="cursor:hand" OnClick='Del_OnClick(<%'RS("Ser_No")%>);'></td>-->
            </tr>
        <%	RS.Movenext
		  Wend%>                                                                                                                   
        </table>
	</form>        
  <%Message()%>
</body>
</html>
<%
	rs.close
	conn.close
%>
<script language="JavaScript"><!--
function Add_OnClick(){
  <%call SubmitJ("add")%>
}
function Update_OnClick(ser_No){  
  document.form.ser_no.value=ser_No;
  <%call SubmitJ("update")%>
}
function Del_OnClick(ser_No){
  document.form.ser_no.value=ser_No;
  <%call SubmitJ("delete")%>
}
--></script>