<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/menu.inc" -->
<% Unitid = Request.QueryString("Unitid") %>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>收件人查詢</title>
<link rel="stylesheet" href="../PageStyle/system1.css">
</head>
<body leftmargin="0">
<center>
  <table border="0" cellspacing="0" cellpadding="0" width="95%" height="35">
    <tr> 
      <td width="37"><img border="0" src="../PageStyle/<%=APCatImageName%>"></td>
      <td background="../PageStyle/banner_a001_b.gif"> 
        <table class="tbc12-c1">
          <tr> 
            <td class="tbc12-c1"><%=Title%></td>
            <td><img border="0" src="../PageStyle/banner_a001-3.gif" hspace="2" width="24" height="18"></td>
            <td class="tbc12-c3">收件人查詢</td>
          </tr>
        </table>
      </td>
      <td background="../PageStyle/banner_a001-5.gif" width="70"><img border="0" src="../PageStyle/banner_a001-4.gif" width="51" height="35"></td>
      <td background="../PageStyle/banner_a001-5.gif" align="right"> 
      <table border="0" cellpadding="0" cellspacing="0">
        <tr class="MenuBody">
   <!-- Menu 開始 --> 
          <% if (HTProgRight and 2)=2 then %><td <% MenuHTML "menu1" %>><a href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>">取消發信</a></td><% End IF %>
   <!-- Menu 結束 -->  
        </tr>
      </table>
      </td>
      <td width="2"><img border="0" src="../PageStyle/banner_a001-6.gif" width="2" height="35"></td>
    </tr>
  </table>
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <tr>
        <td class="c12-1">※ 未選擇條件則顯示所有客戶清單</td>                                                        
    </tr>
  </table>
<form method="POST" action="BulletinEMailList.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>" name="selectform">
  <table border="0" width="60%" cellspacing="1" cellpadding="3">
    <tr>
      <td width="30%" align="right" class="tablebg-c-001">地區：</td>
      <td width="70%" class="tablebg-c-005-3"><Select name="nfx_Area" size=1>
      			<option value="">請選擇</option>
      			<%SQL="Select CC.autoID,CC.C_ITem from ColDefine AS CD Inner Join " & _
      				"ColContain AS CC ON CD.ColID=CC.ColID where CD.ColID=5 Order By CC.SortNo"
      			  SET RSC=conn.execute(SQL)
      			  while not RSC.EOF%>
      				<option value=<%=RSC(0)%>><%=RSC(1)%></option>
      			<%	RSC.movenext
      			  wend
      			%>
      			</select></td>
    </tr>
    <tr>
      <td align="right" class="tablebg-c-001">類別：</td>
      <td class="tablebg-c-005-3"><% SQL="Select CC.autoID,CC.C_ITem from ColDefine AS CD Inner Join " & _
      				"ColContain AS CC ON CD.ColID=CC.ColID where CD.ColID=6 Order By CC.SortNo"
      			  SET RS=conn.execute(SQL)
      			  i=-1
                  while Not rs.eof                 
                  	i=i+1
                  %>
                 <input type=checkbox name="bfx_CompType" value=<%=rs(0)%>><span id=CompType<%=i%>><%= rs(1)%></span>
                <%      rs.movenext
                   wend%></td>
    </tr>
    <tr>
      <td align="right" class="tablebg-c-001">國別：</td>
      <td class="tablebg-c-005-3"><Select name="nfx_Country" size=1>
      			<option value="">請選擇</option>
      			<%SQL="Select CC.autoID,CC.C_ITem from ColDefine AS CD Inner Join " & _
      				"ColContain AS CC ON CD.ColID=CC.ColID where CD.ColID=7 Order By CC.SortNo"
      			  SET RSC=conn.execute(SQL)
      			  while not RSC.EOF%>
      				<option value=<%=RSC(0)%>><%=RSC(1)%></option>
      			<%	RSC.movenext
      			  wend
      			%>
      			</select></td>
    </tr>
    <tr>
      <td align="right" class="tablebg-c-001">業務人員：</td>
      <td class="tablebg-c-005-3"><Select name="nfx_PUserID" size=1>
      			<option value="">請選擇</option>
			<%SQL="Select PUserID,ChName from Personnel where SalesYN='Y'"
			  SET RSS=conn.execute(SQL)
			  if not RSS.EOF then
			  	while not RSS.EOF%>
      				<option value=<%=RSS(0)%>><%=RSS(1)%></option>
<%			  		RSS.movenext
			  	wend
			  end if%>
      			</select></td>
    </tr>
    <tr>
      <td align="right" class="tablebg-c-001">等級：</td>
      <td class="tablebg-c-005-3"><Select name="nfx_EmPower" size=1>
      			<option value="">請選擇</option>
      			<option value=100>A級</option>
      			<option value=99>B級</option>
      			<option value=98>C級</option>
      			</select></td>
    </tr>
    <tr>
      <td align="right" colspan="2"><img border="0" src="../PageStyle/select.gif" OnClick="VBS: selectform.submit" style="cursor:hand;"><img border="0" src="../PageStyle/reset.gif" OnClick="VBS: selectform.reset" style="cursor:hand;"></td>
    </tr>
  </table>
</form>

  </center>                      
</body></html>


