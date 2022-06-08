
<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" 
%>

<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<%
 'session("userId")=ieditor
 'IEditor=session("userID") 
  
 'IDept=session("deptID") 
 IEditor="hyweb"
 IDept="0"
 showType="1"
 siteId="1"
 iBaseDSD="7"
 iCTUnit="2199"
 
 SeciBaseDSD="44"
 SeciCTUnit="2200"
 sql = ""

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
'response.write xUpForm("htx_title") 
'response.write xUpForm("textarea") 
'response.write xUpForm("select") 
'response.write xUpForm("htx_title22") 
'response.write xUpForm("value(startDate)") 
'response.write xUpForm("value(endDate)") 
dim ximportant
if  xUpForm("htx_title22")="" then 
	ximportant = 0 'xUpForm("htx_title22")=0
else
	ximportant = xUpForm("htx_title22")
end if
   'xUpForm("htx_title22")="0"
  

if (xUpForm("htx_title")="" or  xUpForm("value(startDate)")="" or xUpForm("value(endDate)")="" ) then

showDoneBox "輸入不完整！"
 '[response.write "<script language='javascript'>alert('輸入不完整！');history.go(-1);</script>"
response.end

end if
if (xUpForm("value(startDate)") > xUpForm("value(endDate)")) then

showDoneBox "時間輸入有誤！"
 '[response.write "<script language='javascript'>alert('輸入不完整！');history.go(-1);</script>"
response.end

end if
if IsNumeric(xUpForm("htx_title22")) or  xUpForm("htx_title22")="" then
 else
 showDoneBox "重要性--資料為數字！"
 'response.write "<script language='javascript'>alert('(重要性--資料為數字)！');history.go(-1);</script>"
response.end
if int(xUpForm("htx_title22")) < 0 or int(xUpForm("htx_title22"))  > 100 then
 showDoneBox "重要性--99為最高！"
 'response.write "<script language='javascript'>alert('(重要性--99為最高)！');history.go(-1);</script>"
response.end

end if



end if

for each form in xup.Form
if form.IsFile  then
					ofname = Form.FileName
					fnExt = ""
					if instrRev(ofname, ".") > 0 then	fnext = mid(ofname, instrRev(ofname, "."))
					tstr = now()
					nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext 	    
					xup.Form(Form.Name).SaveAs apath & nfname, True			
					sql =  " '" & HTUploadPath2 & nfname & "'"
					
				  ' response.write sql
				end if	
		next
	
		if sql="" then
		sql1="INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD] ,[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[xImportant],[xPostDate],[xPostDateEnd] ,[showType] ,[siteId],[xBody])"
        sql1=sql1&"VALUES('"&iBaseDSD&"','"&iCTUnit&"','"&xUpForm("select")&"','"&xUpForm("htx_title") &"','"&IEditor&"','"&iDept&"',"& ximportant &",'"&xUpForm("value(startDate)")&"','"&xUpForm("value(endDate)")&"','"&showType&"','"&siteId&"','"&xUpForm("textarea") &"')"
		sql1 = "set nocount on;" & sql1 & "; select @@IDENTITY as NewID"	
		else
		sql1="INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD] ,[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[xImportant],[xPostDate],[xPostDateEnd] ,[showType],[xImgFile] ,[siteId],[xBody])"
        sql1=sql1&"VALUES('"&iBaseDSD&"','"&iCTUnit&"','"&xUpForm("select")&"','"&xUpForm("htx_title") &"','"&IEditor&"','"&iDept&"',"&ximportant&",'"&xUpForm("value(startDate)")&"','"&xUpForm("value(endDate)")&"','"&showType&"',"&sql&",'"&siteId&"','"&xUpForm("textarea") &"')"
		sql1 = "set nocount on;" & sql1 & "; select @@IDENTITY as NewID"	
		end if
	
		Set rs = conn.Execute(sql1)
		parentIcuitem=rs(0)
			sql6="insert into CuDTx7 ([giCuItem]) VALUES('"&parentIcuitem&"') "
        Set rs11 = conn.Execute(sql6)
		rs.close
		set rs = nothing 
	
		sql2="INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[topCat],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','Y','最新議題','"&IEditor&"','"&IDept&"','A','"&showType&"','"&siteId&"');"
		sql2 = "set nocount on;" & sql2 & "; select @@IDENTITY as NewID"	
		Set rs = conn.Execute(sql2)
		gicuitem=rs(0)
		sql6="insert into CuDTx7 ([giCuItem]) VALUES('"&gicuitem&"') "
		conn.Execute(sql6)
		rs.close
		set rs = nothing 
		sql2i="INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[parentIcuitem])VALUES("&gicuitem&","&parentIcuitem&")"	
		conn.execute(sql2i)
		
		sql3="INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[topCat],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','Y','議題關聯知識文章','"&IEditor&"','"&IDept&"','B','"&showType&"','"&siteId&"');"
		sql3 = "set nocount on;" & sql3 & "; select @@IDENTITY as NewID"	
		Set rs = conn.Execute(sql3)
		gicuitem=rs(0)
		sql6="insert into CuDTx7 ([giCuItem]) VALUES('"&gicuitem&"') "
		conn.Execute(sql6)
		rs.close
		set rs = nothing 
		sql3i="INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[parentIcuitem])VALUES("&gicuitem&","&parentIcuitem&")"	
		conn.execute(sql3i)
		
		sql4="INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[topCat],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','Y','議題關聯知識文章單元順序設定','"&IEditor&"','"&IDept&"','C','"&showType&"','"&siteId&"');"
		sql4 = "set nocount on;" & sql4 & "; select @@IDENTITY as NewID"	
		Set rs = conn.Execute(sql4)
		gicuitem=rs(0)
		sql6="insert into CuDTx7 ([giCuItem]) VALUES('"&gicuitem&"') "
		conn.Execute(sql6)
		rs.close
		set rs = nothing 
		sql4i="INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[parentIcuitem] ,[orderSiteUnit] ,[orderSubject],[orderKnowledgeTank],[orderKnowledgeHome])VALUES("&gicuitem&","&parentIcuitem&",1,2,3,4)"	
		conn.execute(sql4i)
		
		sql5="INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[topCat],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','Y','議題關聯影音','"&IEditor&"','"&IDept&"','D','"&showType&"','"&siteId&"')"
		sql5 = "set nocount on;" & sql5 & "; select @@IDENTITY as NewID"	
		Set rs = conn.Execute(sql5)
		gicuitem=rs(0)
		sql6="insert into CuDTx7 ([giCuItem]) VALUES('"&gicuitem&"') "
		conn.Execute(sql6)
		rs.close
		set rs = nothing 
		sql5i="INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[parentIcuitem])VALUES("&gicuitem&","&parentIcuitem&")"	
		conn.execute(sql5i)
	
				

        '資源推薦的超連結-bob
        sql5 = " declare @NewIDENTITY bigint "
		sql5 = sql5 & "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[topCat],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','Y','資源推薦的超連結','"&IEditor&"','"&IDept&"','E','"&showType&"','"&siteId&"')"
		sql5 = sql5 & " set @NewIDENTITY = @@IDENTITY "
		sql5 = sql5 & " insert into CuDTx7 ([giCuItem]) VALUES( @NewIDENTITY )"	
		sql5 = sql5 & " INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[parentIcuitem])VALUES(@NewIDENTITY ,"&parentIcuitem&")"	
		conn.Execute(sql5)				
		
		'使用者參與討論或分享心得 added by Joey
		sql6 = " declare @NewIDENTITY bigint "
		sql6 = sql6 & "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[topCat],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','Y','使用者參與討論或分享心得','"&IEditor&"','"&IDept&"','F','"&showType&"','"&siteId&"')"
		sql6 = sql6 & " set @NewIDENTITY = @@IDENTITY "
		sql6 = sql6 & " insert into CuDTx7 ([giCuItem]) VALUES( @NewIDENTITY )"	
		sql6 = sql6 & " INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[parentIcuitem])VALUES(@NewIDENTITY ,"&parentIcuitem&")"	
		conn.Execute(sql6)
		
		'sql2=sql2&"INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','y','議題關聯知識文章','"&IEditor&"','"&IDept&"','"&showType&"','"&siteId&"');"
		'sql2=sql2&"INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[iDept],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','議題關聯知識文章單元順序設定','"&IEditor&"','"&IDept&"','"&showType&"','"&siteId&"');"
		'sql2=sql2&"INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[fCTUPublic],[sTitle],[iEditor],[iDept],[showType],[siteId])VALUES('"&SeciBaseDSD&"','"&SeciCTUnit&"','y','議題關聯影音','"&IEditor&"','"&IDept&"','"&showType&"','"&siteId&"')"
		'conn.execute(sql2)
        'response.write sql2
		
	    response.redirect "index.asp"
		
	
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
              	history.go(-1)
			</script>
    </body>
  </html>
<% End sub %>