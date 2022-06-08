<%@ CodePage = 65001 %>
<%
	Response.charset="utf-8"
%>
<!--#Include virtual = "/inc/client.inc" -->
<%
	'新優質農家 產銷專欄
	sql1 = "SELECT * FROM CuDTGeneric WHERE iCTUnit = '126' and topCat ='01' order by iCUItem"
	set rs1 = conn.execute(sql1)

	while not rs1.eof
		sql_up = "UPDATE CuDTGeneric SET NewtopCat ='02' WHERE iCUItem = '"& rs1("iCUItem") &"'"
		conn.execute(sql_up)
		rs1.movenext
	wend
	
	sql2 = "SELECT * FROM CuDTGeneric WHERE iCTUnit = '126' and topCat ='02' order by iCUItem"
	set rs2 = conn.execute(sql2)
	while not rs2.eof
		sql_up2 = "UPDATE CuDTGeneric SET iCTUnit = '2766' , NewtopCat ='01' WHERE iCUItem = '"& rs2("iCUItem") &"'"
		conn.execute(sql_up2)
		rs2.movenext
	wend
	
	sql3 = "SELECT * FROM CuDTGeneric WHERE iCTUnit = '126' and topCat ='03' order by iCUItem"
	set rs3 = conn.execute(sql3)
	while not rs3.eof
		sql_up3 = "UPDATE CuDTGeneric SET NewtopCat ='01' WHERE iCUItem = '"& rs3("iCUItem") &"'"
		conn.execute(sql_up3)
		rs3.movenext
	wend
	
	sql4 = "SELECT * FROM CuDTGeneric WHERE iCTUnit = '126' and topCat ='04' order by iCUItem"
	set rs4 = conn.execute(sql4)
	while not rs4.eof
		sql_up4 = "UPDATE CuDTGeneric SET iCTUnit = '2766' , NewtopCat ='02' WHERE iCUItem = '"& rs4("iCUItem") &"'"
		conn.execute(sql_up4)
		rs4.movenext
	wend
	
	sql5 = "SELECT * FROM CuDTGeneric WHERE iCTUnit = '126' and topCat ='05' order by iCUItem"
	set rs5 = conn.execute(sql5)
	while not rs5.eof
		sql_up5 = "UPDATE CuDTGeneric SET iCTUnit = '2766' , NewtopCat ='03' WHERE iCUItem = '"& rs5("iCUItem") &"'"
		conn.execute(sql_up5)
		rs5.movenext
	wend
	
	'農業與生活
	sql6 = "SELECT * FROM CuDTGeneric WHERE iCTUnit = '135' order by iCUItem"
	set rs6 = conn.execute(sql6)
	while not rs6.eof
		sql_up6 = "UPDATE CuDTGeneric SET topCat ='02' WHERE iCUItem = '"& rs6("iCUItem") &"'"
		conn.execute(sql_up6)
		rs6.movenext
	wend
	'主題報導專欄
	sql7 = "SELECT * FROM CuDTGeneric WHERE iCTUnit = '299' order by iCUItem"
	set rs7 = conn.execute(sql7)
	while not rs7.eof
		sql_up7 = "UPDATE CuDTGeneric SET iCTUnit = '2766' , NewtopCat ='04' WHERE iCUItem = '"& rs7("iCUItem") &"'"
		conn.execute(sql_up7)
		rs7.movenext
	wend
	response.write "轉資料完成"
%>