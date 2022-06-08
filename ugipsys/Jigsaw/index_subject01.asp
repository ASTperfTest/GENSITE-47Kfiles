<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->

<!--#include virtual = "/inc/client.inc" -->
<%
 'session("userId")=ieditor

 IEditor="hyweb"
 IDept="0"
 showType="1"
 siteId="1"
 iBaseDSD="7"
 iCTUnit="2199"
'if (1=1) then
 'response.write "<script language='javascript'>alert('輸入不完整！');history.go(-1);</script>"
'response.end

'end if
sql=""


function xUpForm(xvar)
		xUpForm = xup.form(xvar)
end function
HTUploadPath = "/public/data/jigsaw/"
HTUploadPath2 = "jigsaw/"
'response.write HTUploadPath
 apath = server.mappath(HTUploadPath) & "\"
Set xup = Server.CreateObject("TABS.Upload")
		xup.codepage = 65001
		xup.Start apath
'response.write xUpForm("sTitle") 

dim ximportant
	
if  xUpForm("htx_title22")="" then  ximportant=0

if xUpForm("sTitle")="" then
showDoneBox1 "請輸入專區標題！"
response.end

end if
 

if  xUpForm("htx_title22")="" then
  else	
	   
		if IsNumeric(xUpForm("htx_title22"))  then
		 else
		 showDoneBox1 "重要性--資料為數字！"
		 'response.write "<script language='javascript'>alert('(重要性--資料為數字)！');history.go(-1);</script>"
		response.end
		end if
end if
if xUpForm("htx_title22")="" then
else
		if int(xUpForm("htx_title22")) < 0 or int(xUpForm("htx_title22"))  > 100  then
		 showDoneBox1 "重要性--99為最高！"
		 'response.write "<script language='javascript'>alert('(重要性--99為最高)！');history.go(-1);</script>"
		response.end

		end if
end if
'response.write xUpForm("textarea") 
'response.write xUpForm("select") 
'response.write xUpForm("htx_title22") 
'response.write xUpForm("value(startDate)") 
'response.write xUpForm("value(endDate)") 
for each form in xup.Form
if form.IsFile  then
					ofname = Form.FileName
					fnExt = ""
					if instrRev(ofname, ".") > 0 then	fnext = mid(ofname, instrRev(ofname, "."))
					tstr = now()
					nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext 	    
					xup.Form(Form.Name).SaveAs apath & nfname, True			
					sql = ","&sql & form.Name & " = '" & HTUploadPath2 & nfname & "'"
					
				   'response.write sql
				end if	
		next
		


	'sql1="UPDATE [mGIPcoanew].[dbo].[CuDTGeneric]SET [fCTUPublic] = '"&xUpForm("select")&"',[sTitle] = '"&xUpForm("sTitle")&"' ,[xImportant] = '"&xUpForm("htx_title22")&"',"
	if  xUpForm("htx_title22")="" then
	sql1="UPDATE [mGIPcoanew].[dbo].[CuDTGeneric]SET [fCTUPublic] = '"&xUpForm("select")&"',[sTitle] = '"&xUpForm("sTitle")&"' ,[xImportant] = "&ximportant&",[xBody] ='"&xUpForm("textarea")&"'"&sql&"  where iCUItem='"&xUpForm("iCUItem")&"'"
    else
	sql1="UPDATE [mGIPcoanew].[dbo].[CuDTGeneric]SET [fCTUPublic] = '"&xUpForm("select")&"',[sTitle] = '"&xUpForm("sTitle")&"' ,[xImportant] = "&xUpForm("htx_title22")&",[xBody] ='"&xUpForm("textarea")&"'"&sql&"  where iCUItem='"&xUpForm("iCUItem")&"'"
    end if

	conn.execute(sql1)
	'response.redirect "index.asp"
	showDoneBox "編修成功！"
	
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
              window.location.href="index.asp"
			</script>
    </body>
  </html>
<% End sub %>
 <% Sub showDoneBox1(lMsg) %>
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
             history.go(-1)
			</script>
    </body>
  </html>
<% End sub %>