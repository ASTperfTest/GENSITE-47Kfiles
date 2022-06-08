<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
'===> 設定共有幾個欄位【請從 0 開始計算】
	rowsno = 1
  
'===> 設定說明內容
	ReadmeText = "※ 檢視請點選「"& Subject &"」"

'===> 設定 SQL
	DefOrderList = ""

    If Request.Form("SelectKey") <> "" or (Request.Form("xfn_Begindate") <> "" and Request.Form("xfn_EndDate") <> "") Then
    SetSQL = "SELECT DISTINCT DataUnit.UnitID, DataUnit.Subject, DataUnit.BeginDate, DataUnit.EndDate "&_
			 "FROM DataUnit RIGHT OUTER JOIN "&_
      		 "DataContent ON DataUnit.UnitID = DataContent.UnitID Where DataUnit.DataType = N'"& DataType &"' And DataUnit.Language = N'"& Language &"'"
	End If
	
     If Request.Form("xfn_Begindate") <> "" and Request.Form("xfn_EndDate") <> "" Then
     	SetSQL = SetSQL & " And (('"& request.form("xfn_Enddate") &"' between DataUnit.BeginDate and DataUnit.EndDate) or ('"& request.form("xfn_Begindate") &"' between DataUnit.BeginDate and DataUnit.EndDate) or (DataUnit.BeginDate between N'"& request.form("xfn_Begindate") &"' and N'"& request.form("xfn_Enddate") &"'))"
     End If 		     
	 
	 If Request.Form("SelectKey") <> "" Then
	  SetSQL = SetSQL & " And (DataContent.Content Like N'%"& Request.Form("SelectKey") &"%' or DataUnit.Subject Like N'%"& Request.Form("SelectKey") &"%')"
	 End If
	 
'===> 欄位排序設定
	OrderField = "" 
	OrderName = ""

     
%>
<!--#Include file = "PageList_Head.inc" --> 
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<title>資料清單 - 進階查詢</title>
</head>
<body>
<object data=../inc/calendar.htm id=calendar type=text/x-scriptlet width=245 height=160 style="position: absolute; top: 30; left: 351; visibility: hidden"></object>
<center>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【進階查詢】</font></td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
          <% if (HTProgRight and 2)=2 then %><a href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>">清單</a><% End IF %>
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    


<br>
<form method="POST" action="BulletinOldList.asp?Language=<%=Language%>&DataType=<%=DataType%>" name="BoardForm">
  <table border="0" cellspacing="0" cellpadding="3" class="whitetablebg">
    <tr>
      <td>公佈期間範圍：</td>
      <td>
        <input type=text name=xfn_Begindate size=10 readonly value="" style="cursor:hand;" onclick="VBScript: btdate 1">
      ～<input type=text name=xfn_Enddate size=10 readonly value="" style="cursor:hand;" onclick="VBScript: btdate 2"></td>
      <td width="10">&nbsp;</td>
      <td>關鍵字：</td>
      <td><input type="text" name="SelectKey" size="15"></td>
      <td><input type="button" value="查詢" OnClick="datacheck()"></td>
    </tr>
  </table>
</form>
<% If SQL <> "" Then %>
<form name="lform" method="POST">    
<!--#Include file = "PageList_body.inc" -->
            <div id="sortTable">                                                
              <table border="0" width="95%" cellpadding="3" cellspacing="1" RULES="GROUPS" FRAME="HSIDES" id="tbx" class="bluetable">                                                
                <THEAD>                                                          
                <tr style="cursor:hand;" class="lightbluetable">                                                          
                  <th id="tbfs0" width="25%">公佈期間&nbsp;<img ID="xtbfs0" border="0" src="../images/i_down.gif"></th>                                                        
                  <th id="tbfs1" width="75%"><%=Subject%>&nbsp;<img ID="xtbfs1" border="0" src="../images/i_down.gif"></th>                                                        
                </tr>                                       
                </THEAD>
             <% for gxo=1 to SetPageCon                                                               
			     if not RS.eof then %>                                      
                <tr class="whitetablebg" id=tbr<%=gxo%>>                                                          
                  <td><%=rs("BeginDate")%> ~ <%=rs("EndDate")%></td>                                                        
                  <td><a href="BulletinView.asp?Language=<%=Language%>&amp;DataType=<%=DataType%>&amp;UnitID=<%=rs("UnitID")%>"><%=rs("Subject")%></a></td>                                                        
                </tr>                                       
			 <%  RS.moveNext                                                             
			     End IF                                                             
			    Next %>                                                             
			  </table>
			</div>
</form>
<% End If %>

    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Best viewed with Internet Explorer 5.5+ , Screen Size set to 800x600.</td>                                           
  </tr> 
</table> 
</center>
</body></html>
<!--#Include file = "PageList_End.inc" --> 
<script language="vbscript">
Sub datacheck()
  msg2="請輸入查詢條件！"  

  If BoardForm.xfn_BeginDate.value = Empty And BoardForm.xfn_EndDate.value = Empty And BoardForm.SelectKey.value = Empty Then
     MsgBox msg2, 16, "Sorry!"    
     BoardForm.SelectKey.focus    
     Exit Sub
  End IF

  msg3 = "請輸入「公佈起始時間」！"  
  msg4 = "請輸入「公佈結束時間」！"  
  msg5 = "「公佈結束時間」早於「公佈起始時間」，請重新輸入 !"
  If BoardForm.SelectKey.value = Empty Then  
   If BoardForm.xfn_BeginDate.value = Empty Then
     MsgBox msg3, 16, "Sorry!"
     BoardForm.xfn_BeginDate.focus
     Exit Sub
   ElseIf BoardForm.xfn_EndDate.value = Empty Then
     MsgBox msg4, 16, "Sorry!"
     BoardForm.xfn_EndDate.focus
     Exit Sub
   ElseIf CDate(BoardForm.xfn_BeginDate.value) > CDate(BoardForm.xfn_EndDate.value) Then    
     MsgBox msg5, 16, "Sorry!"    
     BoardForm.xfn_EndDate.focus    
     Exit Sub
   End if
  End If
  
  BoardForm.Submit
End Sub

dim CanTarget
sub btdate(n)             
 If document.all.calendar.style.visibility="" Then                
   document.all.calendar.style.visibility="hidden"             
 Else             
   document.all.calendar.style.visibility=""              
 End If                   
 CanTarget=n     
end sub                
                               
sub calendar_onscriptletevent(n,o)                
  document.all.calendar.style.visibility="hidden"                
  select case CanTarget     
     case 1     
          document.all.xfn_Begindate.value=n
     case 2     
          document.all.xfn_Enddate.value=n
  end select     
end sub 
</script>
