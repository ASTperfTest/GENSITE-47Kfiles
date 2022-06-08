<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/menu.inc" -->
<% unitid = Request.QueryString("unitid") %>
<%
 ckno = 0

	SQL = "Select C.ContactEMail, C.CompTypeStr,C.Area,C.CompType,C.Country,C.CompanyID,C.SEName,C.ContactName,C.AddDate,P.PUserID,P.ChName AS SalesName," & _
     	   "(Select C_Item from ColContain where autoID=C.Area) AS AName," & _ 
     	   "(Select C_Item from ColContain where autoID=C.Country) AS CName" & _  
     	   " From Customer AS C Left Join Personnel AS P ON C.PUserID=P.PUserID " & _
     	   "Where Mark='1'" 
	If Request.Form("nfx_Area") <> "" Then SQL = SQL & " And Area = "& Request.Form ("nfx_Area")
	If Request.Form("nfx_Country") <> "" Then SQL = SQL & " And Country = "& Request.Form ("nfx_Country")
	If Request.Form("nfx_PUserID") <> "" Then SQL = SQL & " And C.PUserID = "& Request.Form ("nfx_PUserID")
	If Request.Form("nfx_EmPower") <> "" Then SQL = SQL & " And EmPower = "& Request.Form ("nfx_EmPower")
	
                    		if Request.Form("bfx_CompType") <> "" then
                    			xstr = Request.Form("bfx_CompType")
                    			xpos = Instr(xstr,",")
                    			SQLstr = " AND ("
                    			if xpos= 0 then
            					 SQL = SQL & " AND C.CompType LIKE N'%" & Request.Form("bfx_CompType") & "%'"            					                    			
                    			else
                    			   while xpos<>0 
                    				xxstr=left(xstr,xpos-1)
            					    SQLstr = SQLstr & "C.CompType LIKE N'%" & xxstr & "%' OR "
            					    xstr=mid(xstr,xpos+1)
            					    xpos=Instr(xstr,",")
            					    if xpos=0 then SQLstr = SQLstr & "C.CompType LIKE N'%" & xstr & "%')"            					
            				       wend
            				   SQL = SQL & SQLstr
            				   End IF
            				end if
    Set rs = conn.Execute(SQL)
 	   
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>收件人查詢清單</title>
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
            <td class="tbc12-c3">收件人清單</td>
          </tr>
        </table>
      </td>
      <td background="../PageStyle/banner_a001-5.gif" width="70"><img border="0" src="../PageStyle/banner_a001-4.gif" width="51" height="35"></td>
      <td background="../PageStyle/banner_a001-5.gif" align="right"> 
      <table border="0" cellpadding="0" cellspacing="0">
        <tr class="MenuBody">
   <!-- Menu 開始 --> 
          <% if (HTProgRight and 2)=2 then %><td <% MenuHTML "menu1" %>><a href="BulletinEMail.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>">查詢</a></td><% End IF %>
          <% if (HTProgRight and 2)=2 then %><td <% MenuHTML "menu2" %>><a href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>">取消發信</a></td><% End IF %> 
   <!-- Menu 結束 -->  
        </tr>
      </table>
      </td>
      <td width="2"><img border="0" src="../PageStyle/banner_a001-6.gif" width="2" height="35"></td>
    </tr>
  </table>
<% If Not rs.EOF Then %>
<form name="reg" method="POST" action="BulletinEMailSet.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>">
             <table border="0" width="95%" cellpadding="3" cellspacing="1">                                                
                <THEAD>                                                          
                <tr>                                                          
                  <th class="tablebg-c-005-2" width="7%"><img id="xtbfs5" border="0" src="../PageStyle/all.gif" OnClick="CheckBoxed()" style="cursor:hand;"></th>                                                        
                  <th id="tbfs0" class="tablebg-c-005-2" width="13%">地區</th>                                                        
                  <th id="tbfs1" class="tablebg-c-005-2" width="10%">國別</th>                                                        
                  <th id="tbfs2" class="tablebg-c-005-2" width="30%">公司簡稱</th>                                                        
                  <th id="tbfs3" class="tablebg-c-005-2" width="20%">聯絡人</th>                                                        
                  <th id="tbfs4" class="tablebg-c-005-2" width="20%">業務人員</th>                                                       
                </tr>                                       
                </THEAD>                                      
                <TBODY>
<% Do While Not RS.EOF 
    ckno = ckno + 1 %>                                                      
  <tr class="tablebg-c-005-3"> 
    <td align="center"><input type="checkbox" name="MailCk" value="<%=rs("ContactEMail")%>"></td>
    <td><%=rs("AName")%></td>
    <td><%=rs("CName")%></td>
    <td><%=RS("SEName")%></td>
    <td><%=rs("ContactName")%></td>
    <td><%=rs("SalesName")%></td>
  </tr>
<% RS.MoveNext
   Loop %>  
  <tr> 
    <td colspan="6" align="right"><input type="hidden" name="ck" value="N"><img border="0" src="../PageStyle/enter.gif" OnClick="sbumitck()" style="cursor:hand;"><img border="0" src="../PageStyle/reset.gif" OnClick="VBS: reg.reset" style="cursor:hand;"></td>
  </tr>
  </TBody>
</table>
</form>
<% Else %>
  <p>　</p>
  <p>　</p>
  <p><font color="#FF0000" size="2">查無資料</font></p>
  <p>　</p>
  <p><font size="2"><a href="BulletinEMail.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>">重設查詢</a></font>
<% End If %>
  </center>                      
</p>
</body></html>
<script language=vbs>
Sub CheckBoxed()
 ckno = <%=ckno%>
 If ckno <> 0 Then
  If document.all.ck.value = "N" Then
    If ckno = 1 Then
     document.all.MailCk.checked = True
    Else
     For xno = 0 to ckno-1
      document.all.MailCk(xno).checked = True
     Next
    End If
     document.all.ck.value = "Y"
     document.all.xtbfs5.src = "../PageStyle/not.gif"
  Else
    If ckno = 1 Then
     document.all.MailCk.checked = False
    Else
     For xno = 0 to ckno-1
      document.all.MailCk(xno).checked = False
     Next
    End If
     document.all.ck.value = "N"
     document.all.xtbfs5.src = "../PageStyle/all.gif" 
  End If
 End IF
End Sub

function sbumitck()               
<% if ckno>1 then %>               
  for i=0 to reg.MailCk.length-1               
    if reg.MailCk(i).checked then               
      radiochk="YES"               
      exit for               
    end if               
  next               
<% else %>               
  if reg.MailCk.checked then               
    radiochk="YES"               
  end if               
<% end if %>               
  if radiochk<>"YES" then               
    alert("尚未選定!!")               
  else         
    chky=msgbox("　確定要寄出嗎？"& vbcrlf , 48+1, "請注意！！")
    if chky=vbok then             
    reg.submit               
  end if               
 end if               
end function               

</script>

