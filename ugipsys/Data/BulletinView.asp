<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<% UnitID = Request.QueryString("UnitID") %>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>資料檢視</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<center>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【檢視】</font></td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
          <% if (HTProgRight and 8)=8 then %><a href="BulletinEdit.asp?Language=<%=Language%>&amp;DataType=<%=DataType%>&UnitID=<%=UnitID%>">編修</a><% End IF %>
          <% if (HTProgRight and 2)=2 then %>　<a href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>">查詢</a><% End IF %>
   　   </td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    

<br>
<%
 If CatDecide = "Y" Then
  SQL = "Select DataCat.CatName, DataCat.CatID, DataUnit.* From DataUnit, DataCat Where DataCat.CatID = DataUnit.CatID And UnitID = "& UnitID
 Else
  SQL = "Select * From DataUnit Where UnitID = "& UnitID
 End If
  SET RS = conn.execute(SQL)
   EditUser = rs("EditUserID")
   EditDate = rs("EditDate")
%>
<table border="0" width="580" cellspacing="1" cellpadding="3" class="bluetable">
<% If CatDecide = "Y" Then %>
  <tr><td width="100" class="lightbluetable">類別</td>
      <td width="480" class="whitetablebg"><%=rs("CatName")%></td></tr>
<% End IF %>
  <tr><td width="100" class="lightbluetable"><%=Subject%></td>
      <td width="480" class="whitetablebg"><%=rs("Subject")%></td></tr>
<% If Not IsNull(Extend_1) Then %>
  <tr><td class="tablebg-c-001-1"><%=Extend_1 &"：http://"& rs("Extend_1")%></td></tr>
<% End IF %>
<% If DateDecide = "Y" Then %>
  <tr><td width="100" class="lightbluetable">公佈時間</td>
      <td width="480" class="whitetablebg"><%=rs("BeginDate")%>～<%=rs("EndDate")%></td></tr>
<% End IF %>
</table><br>
<table border="0" width="90%" cellspacing="0" cellpadding="3">
  <tr>
    <td width="100%" class="whitetablebg">
<% 	SQLCom = "SELECT * FROM DataContent Where UnitID = "& UnitID &" Order By Position"
	SET RSCom = conn.execute(SQLCom)
	 If Not rscom.EOF Then
	  DIVCount = 0
	  Do while not rscom.EOF
	   DIVCount = DIVCount + 1
	   ImageHTELSrc = ""
	   ClientCk = "N"
	   comm = rscom("Content")
	   ncomm = message(comm)
	   If Not (rscom("ImageFile") = "" or IsNull(rscom("ImageFile"))) Then ImageHTELSrc = "<img src=../Public/Data/"& rscom("ImageFile") &" border=0 align="& rscom("ImageWay") &" id=conimg"& rscom("Position") &" alt=""第 "& rscom("Position") &" 段圖片"">"
	   response.write ImageHTELSrc & ncomm &"<p>"& vbcrlf
	  rscom.movenext
      loop
     End If %>
    </td>
  </tr>
</table>

    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Best viewed with Internet Explorer 5.5+ , Screen Size set to 800x600.</td>                                          
  </tr> 
</table> 
</center>
</body></html>
<%
function message(tempstr)
  outstring = ""
  while len(tempstr) > 0
    pos=instr(tempstr, chr(13)&chr(10))
    if pos = 0 then
      outstring = outstring & tempstr & "<br>"
      tempstr = ""
    else
      outstring = outstring & left(tempstr, pos-1) & "<br>"
      tempstr=mid(tempstr, pos+2)
    end if
  wend
  message = outstring
end function
%>