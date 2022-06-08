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

orderArticle = "1"
IEditor = "hyweb"
IDept = "0"
showType = "1"
siteId = "1"
iBaseDSD = "44"
iCTUnit = "2201"
CtRootId = 0

insert1=0   '新增的超連結，沒有原始文章 

select case request("AddLinkAction")
    case "AddLink"  '加入資源推薦link
        
            parentIcuitem=request("gicuitem")    'Id : 資源推薦的超連結        
            title = request("title")            
            if left(ucase(request("url")), 7) <> "HTTP://" then
                url = "http://" & request("url")
            else
                Url  = request("url")
            end if        

			sql2 = "declare @newIDENTITY bigint"
			sql2 = sql2 & vbcrlf & "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[dEditDate],[iDept],[showType],[siteId]) "
			sql2 = sql2 & vbcrlf & "VALUES(" & iBaseDSD & ", " & iCTUnit & ", '" & title & "', '" & IEditor & "', GETDATE(), '" & IDept & "', '" & showType & "', '" & siteId & "') "
			sql2 = sql2 & vbcrlf & "set @newIDENTITY = @@IDENTITY "
			sql2 = sql2 & vbcrlf & ""
			sql2 = sql2 & vbcrlf & "INSERT INTO CuDTx7 ([giCuItem]) VALUES(@newIDENTITY)"
			sql2 = sql2 & vbcrlf & ""
		    sql2 = sql2 & vbcrlf & "INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[CtRootId],[CtUnitId],[parentIcuitem],[ArticleId],[Status],[orderArticle],[path]) "
			sql2 = sql2 & vbcrlf & "VALUES(@newIDENTITY, " & CtRootId & ", 1, " & parentIcuitem & ", " & insert1 & ", 'Y', " & orderArticle & ", '" & Url & "')"						
			
			conn.execute sql2     
        'response.clear()
        'response.write sql2
        'RESPONSE.END
    case else       '編修存檔
        sql2="SELECT KnowledgeJigsaw.gicuitem  FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem WHERE (KnowledgeJigsaw.Status = 'y') and ( KnowledgeJigsaw.parentIcuitem='"&request("gicuitem")&"')"

        set rs2=conn.Execute(sql2)
        while not rs2.eof

        if request(rs2("gicuitem"))<>"" then
        sql4="SELECT  [orderArticle] FROM [mGIPcoanew].[dbo].[KnowledgeJigsaw] where gicuitem='"&rs2("gicuitem")&"'"
        set rs4=conn.Execute(sql4)

        if isnumeric(request(rs2("gicuitem"))) then        
          if int(rs4(0))<> int(request(rs2("gicuitem"))) then
           sql ="UPDATE [mGIPcoanew].[dbo].[KnowledgeJigsaw] SET [orderArticle] ="&request(rs2("gicuitem"))&"  WHERE gicuitem='"&rs2("gicuitem")&"'"
           conn.Execute(sql)
           sql3="UPDATE [mGIPcoanew].[dbo].[CuDTGeneric] SET [dEditDate] =getdate()  WHERE iCUItem ='"&rs2("gicuitem")&"'"
           conn.Execute(sql3)
          end if 
        end if

        end if 
        rs2.movenext
        wend
end select

showDoneBox "編修成功！"
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