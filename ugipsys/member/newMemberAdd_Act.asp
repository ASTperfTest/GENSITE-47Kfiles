<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="newmember"
HTProgPrefix = "newMember"
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include virtual = "/inc/client.inc" -->
<% 

	HTUploadPath = session("Public")
	apath = server.mappath(HTUploadPath) & "\"
	
		Set xup = Server.CreateObject("TABS.Upload")
		xup.codepage = 65001
		xup.Start apath
		
function xUpForm(xvar)
	xUpForm = trim(xup.form(xvar))
end function

account=xUpForm("account")
id=xUpForm("id")

accountsql="SELECT  COUNT(*) AS ACCOUNT from Member where account="&"'" & account &"'"

Set RSreg=conn.execute(accountsql) 

	if  RSreg("ACCOUNT")="1" then
	   response.write "<script>alert('此帳號已註冊');history.back();</script>"
    end if 
if id <> "" then
	idsql="SELECT  COUNT(*) AS id from Member where id="&"'" & id &"'"
	Set RSreg1=conn.execute(idsql) 
	if RSreg1("id")="1" then
	  response.write "<script>alert('此身分證帳號已註冊');history.back();</script>"
	end if
end if
for each form in xup.Form
	Formname=Form.name
	if xUpForm ( "" &  Formname  & "")<>""  and  Formname<>"submit" and Formname<>"password2" and Formname<>"birthYear"and Formname<>"birthMonth"and Formname<>"birthday" and Formname<>"photo" and Formname<>"id_type" then
	
	sql1 = sql1 &  Formname  & ","
	sql2= sql2 &"'"& xUpForm(""&Formname&"")&"'"& ","
	end if
	
	next	
	
sql1=Mid(sql1,1,len(sql1)-1)
sql2=Mid(sql2,1,len(sql2)-1) 




 if xUpForm ("id_type")="1" then
 sql1=sql1 & ","& "id_type1" & "," & "scholarValidate"
 sql2=sql2 & ","&"'"& 1 &"'" & "," & "'Z'"
 end if
  if xUpForm ("id_type")="2" then
  sql1=sql1 & ","& "id_type1" &","& "id_type2" & "," & "scholarValidate"
  sql2=sql2 & ","&"'"& 1 &"'" & ","&"'"& 1 &"'" & "," & "'Y'"
 end if
 
 
if xUpForm ("birthYear")<>"" and xUpForm ("birthMonth")<>""  and xUpForm ("birthday")<>"" then
sql1=sql1 & ","& "birthday"
sql2=sql2 & ","&"'"& xUpForm ("birthYear") &xUpForm ("birthMonth")& xUpForm ("birthday") &"'"
end if 

if  xUpForm ("photo")<>"" then
nfname = xup.Form("photo").FileName
pos = instr(nfname, ".")
nfname = Year(now()) & Month(now()) & day(now()) & Hour(now())& Minute(now()) & Second(now()) &"." & mid(nfname, pos + 1)
route =  "\"& "public" & "\" & nfname
'xup.Form("photo").SaveAs apath & nfname, True
sql1=sql1 & ","& "photo"	
sql2 = sql2 & ","& "'" & route  &"'"
end if 

sql1=sql1 & "," & "createtime" &","& "modifytime" & "," & "logincount" & ","& " status"
sql2=sql2 & ", GETDATE(),GETDATE()," & 0 & "," & "'Y'"



sql = "INSERT INTO Member(" & Mid(sql1,1,len(sql1)) & ") VALUES(" & Mid(sql2,1,len(sql2)) & ")"


conn.execute(sql)
'response.write "<script>alert('此帳號已註冊');history.back();</script>"
response.write "<script>alert('此帳號註冊成功');location.href='newMemberList.asp';</script>"
'response.redirect "newMemberList.asp"


%>
