<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/client.inc" -->

<%


iCUItem=request("iCUItem")
check=request("check")
uncheck=request("uncheck")
check1=session("jigcheck")
a=session("jigcheck1")
b=check
c=uncheck
checkarr = split(b, ";")
uncheckarr = split(c, ";")
'response.write request("gicuitem")

For i = 0 To UBound(checkarr)
   if (InStr(a,checkarr(i))>0) then
	
   else
	   add=checkarr(i)+";"
	   a=a+add
	end if
next

For i = 0 To UBound(uncheckarr)-1
    if (InStr(a,uncheckarr(i))>0) then
	   cut=uncheckarr(i)+";"
       a=Replace(a,cut,"")
    end if
next
session("jigcheck1")=a
update1=split(a, ";")
For i = 0 To UBound(update1)-1
	sql = "UPDATE [mGIPcoanew].[dbo].[KnowledgeJigsaw] SET [Status] ='N' WHERE [gicuitem] ='"&update1(i)&"'"
	conn.Execute(sql)
	'added by Joey,2009/10/12,http://gssjira.gss.com.tw/browse/COAKM-9 , 刪除的同時移除該User的KPI相關數值
	'modified by Joey, 2009/10/26, http://gssjira.gss.com.tw/browse/COAKM-19, 當管理員刪除網友留言後，要將[shareJigsaw] 欄位值-1
	sql2="SELECT iEditor , convert(varchar, dEditDate, 111) as dEditDate FROM CuDTGeneric WHERE iCUItem=" & update1(i)
	set RS = conn.execute(sql2)
	if RS.EOF=false then
		sql3="update MemberGradeShare set MemberGradeShare.shareJigsaw=MemberGradeShare.shareJigsaw-1 where MemberGradeShare.memberId='" & RS("iEditor") & "' and convert(varchar, MemberGradeShare.shareDate, 111)='" & RS("dEditDate") & "'"
		conn.Execute(sql3)
		'response.write "mysql3:" & sql3 &"<br>"
	end if



next
showDoneBox "刪除成功！"
'response.redirect "subjectPubList.asp?iCUItem="&request("iCUItem")
 %>
 <% Sub showDoneBox(lMsg) %>
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
              window.location.href="subjectPubList.asp?iCUItem=<%=request("iCUItem")%>"
			</script>
    </body>
  </html>
<% End sub %>