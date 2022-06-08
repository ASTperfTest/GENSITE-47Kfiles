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


orderSiteUnit= request("orderSiteUnit")
orderSubject= request("orderSubject")
orderKnowledgeTank= request("orderKnowledgeTank")
orderKnowledgeHome= request("orderKnowledgeHome")
check="1234"
if (InStr(check,orderSiteUnit)>0) then
 check=Replace(check,orderSiteUnit,"")
 else
 response.write "<script language='javascript'>alert('輸入不正確！');history.go(-1);</script>"
 response.end
end if
 if (InStr(check,orderSubject)>0) then
 check=Replace(check,orderSubject,"")
 else
 response.write "<script language='javascript'>alert('輸入不正確！');history.go(-1);</script>"
 response.end
end if
 if (InStr(check,orderKnowledgeTank)>0) then
 check=Replace(check,orderKnowledgeTank,"")
 else
 response.write "<script language='javascript'>alert('輸入不正確！');history.go(-1);</script>"
 response.end
end if
 if (InStr(check,orderKnowledgeHome)>0) then
 check=Replace(check,orderKnowledgeHome,"")
 else
 response.write "<script language='javascript'>alert('輸入不正確！');history.go(-1);</script>"
 response.end
end if


 sql="SELECT CuDTGeneric.sTitle as sTitle, KnowledgeJigsaw.parentIcuitem as parentIcuitem, KnowledgeJigsaw.gicuitem as gicuitem ,CuDTGeneric.iCUItem as iCUItem FROM KnowledgeJigsaw INNER JOIN CuDTGeneric ON KnowledgeJigsaw.gicuitem = CuDTGeneric.iCUItem where KnowledgeJigsaw.parentIcuitem='" &request("iCUItem")&"' AND (CuDTGeneric.sTitle <> '議題關聯知識文章單元順序設定')"           

 Set rs= conn.Execute(sql)

 while not rs.eof
sql1="UPDATE [mGIPcoanew].[dbo].[CuDTGeneric] SET  fCTUPublic ='"&request(rs("gicuitem"))&"' WHERE (iCUItem = '"&rs("iCUItem")&"')"

conn.Execute(sql1)

rs.movenext


wend

sql2="SELECT CuDTGeneric.sTitle as sTitle, KnowledgeJigsaw.parentIcuitem as parentIcuitem, KnowledgeJigsaw.gicuitem as gicuitem ,CuDTGeneric.iCUItem as iCUItem FROM KnowledgeJigsaw INNER JOIN CuDTGeneric ON KnowledgeJigsaw.gicuitem = CuDTGeneric.iCUItem where KnowledgeJigsaw.parentIcuitem='" &request("iCUItem")&"' AND (CuDTGeneric.sTitle = '議題關聯知識文章單元順序設定')" 
'response.write sql2
Set rs2= conn.Execute(sql2)

sql3="UPDATE [mGIPcoanew].[dbo].[KnowledgeJigsaw]SET [orderSiteUnit] = '"&orderSiteUnit&"',[orderSubject] = '"&orderSubject&"',[orderKnowledgeTank] = '"&orderKnowledgeTank&"',[orderKnowledgeHome] = '"&orderKnowledgeHome&"' WHERE gicuitem='"&rs2("gicuitem")&"'"

conn.Execute(sql3)

 
response.redirect "index.asp"
response.end
%>