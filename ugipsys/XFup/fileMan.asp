<%@ CodePage = 65001 %>
<%
Set fso = server.CreateObject("Scripting.FileSystemObject")

xPath = request("xPath")
upPath = session("uploadPath") & xPath
if right(upPath,1)<>"/" then	upPath = upPath & "/"
if right(xPath,1)<>"/" then	xPath = xPath & "/"
if left(xPath,1)="/" then	xPath = mid(xPath,2)

Set fldr = fso.GetFolder(server.MapPath(upPath))
 
	%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
	<head>
	</head>
	<body>
	
		<form action="upload.asp" method="post" enctype="multipart/form-data" ID="Form1">
			<table cellpadding="3" cellspacing="1" border="1" width="98%">
			        <tr>
					<td  width="100%" colspan="2"  bgcolor="lightblue" align="center">
					目前路徑：  <%=session("uploadPath")%><%=xpath%>　　　　
					<input type="hidden" name="xxPath" value="<%=xpath%>">
                    子目錄名稱 <input name="newFolder" size=20> <input type="submit" value="新增" border="0" name="send">
					</td>
					
				</tr>
			    <tr>
					<th bgcolor="lightblue" align="center">上 傳 檔 案</th>
					<th bgcolor="lightblue" align="center">檔 案 清 單</th>
				</tr>
				<tr>
				<td valign="top"><table width="100%">
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案1</td>
					<td >
						<input name="photo1" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案2</td>
					<td >
						<input name="photo2" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案3</td>
					<td >
						<input name="photo3" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案4</td>
					<td >
						<input name="photo4" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案5</td>
					<td >
						<input name="photo5" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案6</td>
					<td >
						<input name="photo6" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案7</td>
					<td >
						<input name="photo7" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案8</td>
					<td >
						<input name="photo8" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案9</td>
					<td >
						<input name="photo9" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">檔案0</td>
					<td >
						<input name="photo0" type="file" size="30">
					</td>
				</tr>
				<tr>
					<td colspan="2" align="right">
						<input type="reset" value="重設" border="0"  name="reset" ID="Reset1"> <input type="submit" value="送出" border="0" name="send" id="send">
					</td>
				</tr>
					</table>
				</form>
				</td>
				<td valign="top">
				<form name="reg" method="POST" action="delProcess.asp">
					<input type="hidden" name="xxPath" value="<%=xpath%>">
					<input type="hidden" name="submitTask" value="">
				<table width="100%">
<%
	if fldr.Files.count > 0 OR fldr.subFolders.count > 0 then
%>
				<tr bgcolor="lightblue">
					<td><INPUT type="button" name="delFile" VALUE="刪檔" onClick="VBS:formDelFileSubmit()"></td>
					<td>檔 名</td>
					<td align="right">大 小</td>
				</tr>	
			        <% 
			            for each sf in fldr.Files
			        %>
			        <tr bgcolor="yellow">
			           <td ><input type="checkbox" name="insertFiles" value="<%=sf.name%>" ></td>
                       <td><A target="_nwMof" href="<%=session("uploadPath")%><%=xpath%><%=sf.name%>">
                       <%=sf.name%></A></td>
                        <td align="right"><%=sf.size%></td>
                              	</tr>
                               	<%
                               	    next                               	
                               	 %>
<%	else %>
<tr><th>
<INPUT type="button" name="delDir" VALUE="刪除此目錄" onClick="VBS:formDelDirSubmit()">
</th></tr>
<%	end if %>				
				</table></td></tr>
				
			</table>
		</form>
	</body>
<script language=vbs>
   sub formDelFileSubmit()
         chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除檔案嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	       reg.submitTask.value = "delFiles"
	      reg.Submit
       end If
  end sub
   sub formDelDirSubmit()
         chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除目錄嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	       reg.submitTask.value = "delDir"
	      reg.Submit
       end If
  end sub

</script>
</html>
