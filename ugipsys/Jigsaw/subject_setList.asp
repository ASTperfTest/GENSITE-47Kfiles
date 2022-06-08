<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" 
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<%

'response.write session("jigsql") 
'Set conn = Server.CreateObject("ADODB.Connection")
	'StrCn="Provider=SQLOLEDB;Data Source=10.10.5.127;User ID=hygip;Password=hyweb;Initial Catalog=mGIPcoanew"

'----------HyWeb GIP DB CONNECTION PATCH----------
'	'conn.Open StrCn
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
'conn.ConnectionString = StrCn
'----------HyWeb GIP DB CONNECTION PATCH----------


'----------HyWeb GIP DB CONNECTION PATCH----------
'	'conn.Open
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
'conn.ConnectionString = 
'----------HyWeb GIP DB CONNECTION PATCH----------

if request("clear") = "1" then
	session("jigsql") = ""
end if	
	
sql="SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("iCUItem")&"'"
Set rs = conn.Execute(sql)
tilte=rs("stitle")
sqlG=""

StartDate=Request("value(startDate)") 
EndDate=Request("value(endDate)") 
if session("jigsql")="" then
If Request("CtRootId") = "4" Then		
		Set KMConn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'		KMConn.Open session("KMODBC")
'Set KMConn = Server.CreateObject("HyWebDB3.dbExecute")
KMConn.ConnectionString = session("KMODBC")
KMConn.ConnectionTimeout=0
KMConn.CursorLocation = 3
KMConn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

		
		sql = ""
		sql = sql & " SELECT DISTINCT REPORT.REPORT_ID, CATEGORY.CATEGORY_NAME, CATEGORY.CATEGORY_ID, REPORT.SUBJECT, REPORT.PUBLISHER, REPORT.ONLINE_DATE, "
		sql = sql & " ACTOR_INFO.ACTOR_DETAIL_NAME FROM REPORT INNER JOIN ACTOR_INFO ON REPORT.CREATE_USER = ACTOR_INFO.ACTOR_INFO_ID "
		sql = sql & " INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY ON "
		sql = sql & " CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID AND CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID "
		sql = sql & " INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID "
		sql = sql & " WHERE (REPORT.STATUS = 'PUB') AND (REPORT.ONLINE_DATE < GETDATE()) AND (CATEGORY.DATA_BASE_ID = 'DB020') "
		sql = sql & " AND (RESOURCE_RIGHT.ACTOR_INFO_ID IN('001','002') ) "
		
		If Request("CtNodeName") <> "" Then
			sql = sql & "AND (CATEGORY.CATEGORY_NAME LIKE " & pkDate("%" & Request("CtNodeName") & "%", "") & ") "
		End If	
		
		If Request("sTitle") <> "" Then
			sql = sql & "AND (REPORT.SUBJECT LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "
		End If
		
		If StartDate <> "" And EndDate <> "" Then
			sql = sql & "AND (REPORT.ONLINE_DATE BETWEEN " & pkStr(StartDate, "") & " AND " & pkStr(EndDate, "") & ") "
		ElseIf StartDate <> "" Then
			sql = sql & "AND (REPORT.ONLINE_DATE >= " & pkStr(StartDate, "") & ") "
		ElseIf EndDate <> "" Then
			sql = sql & "AND (REPORT.ONLINE_DATE <= " & pkStr(EndDate, "") & ") "
		End If
		
		If Request("Status") = "Y" Then
			sql = sql & "AND (REPORT.STATUS = 'PUB') "
		Else
			sql = sql & "AND (REPORT.STATUS <> 'PUB') "
		End If
		
		sql = sql & " ORDER BY REPORT.ONLINE_DATE DESC "
		'If Request("DeptId") <> "" Then
		'	sql = sql & "AND (CuDTGeneric.iDept = " & pkStr(Request("DeptId"), "") & ")"
		'End If
		'response.write sql
		session("jigsql")=sql
	ElseIf Request("CtRootId") = "3" Then		
		
		
		sql = ""
		sql = sql & " SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, CatTreeNode.CatName, "
		sql = sql & " CtUnit.CtUnitName, CuDTGeneric.dEditDate, Member.realname FROM CuDTGeneric INNER JOIN CatTreeNode ON "
		sql = sql & " CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON CatTreeNode.CtRootID = CatTreeRoot.CtRootID "
		sql = sql & " INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
		sql = sql & " INNER JOIN KnowledgeForum ON CuDTGeneric.icuitem = KnowledgeForum.gicuitem "
		sql = sql & " WHERE (CuDTGeneric.siteId = N'3') AND (CatTreeRoot.inUse = 'Y') AND (CuDTGeneric.fCTUPublic = 'Y') "
		sql = sql & " AND (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') AND (CatTreeRoot.inUse = 'Y') AND (CatTreeNode.CtUnitID = 932) "
		sql = sql & " AND (KnowledgeForum.Status <> 'D')"
		
		If Request("sTitle") <> "" Then
			sql = sql & "AND (CuDTGeneric.sTitle LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "
		End If
		
		If StartDate <> "" And EndDate <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate BETWEEN " & pkStr(StartDate, "") & " AND " & pkStr(EndDate, "") & ") "
		ElseIf StartDate <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate >= " & pkStr(StartDate, "") & ") "
		ElseIf EndDate <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate <= " & pkStr(EndDate, "") & ") "
		End If
		
		If Request("Status") = "Y" Then
			sql = sql & "AND (CuDTGeneric.fCTUPublic = 'Y') "
		Else
			sql = sql & "AND (CuDTGeneric.fCTUPublic <> 'Y') "
		End If
		sql = sql & " ORDER BY CuDtGeneric.dEditDate DESC "

        session("jigsql")=sql		
	Else
		
		sql = ""	
		sql = sql & "SELECT DISTINCT CatTreeNode.CtRootID, CatTreeNode.CtNodeID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, "
		sql = sql & "CatTreeNode.CatName, CtUnit.CtUnitName, InfoUser.UserName, CuDTGeneric.dEditDate, Dept.deptName FROM CuDTGeneric INNER JOIN "
		sql = sql & "CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON "
		sql = sql & "CatTreeNode.CtRootID = CatTreeRoot.CtRootID INNER JOIN CtUnit ON "
		sql = sql & "CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN InfoUser ON "
		sql = sql & "CuDTGeneric.iEditor = InfoUser.UserID INNER JOIN Dept ON CuDTGeneric.iDept = Dept.deptID "
		sql = sql & "WHERE (CatTreeRoot.inUse = 'Y') AND (CtUnit.inUse = 'Y') "
		sql = sql & "AND (CatTreeRoot.inUse = 'Y') AND (Dept.inUse = 'Y')"
		
		If Request("CtRootId") <> "" Then		
			If Request("CtRootId") = "1" Then			
				sql = sql & "AND (CuDtGeneric.siteId = " & pkStr(Request("CtRootId"), "") & ") AND (CatTreeNode.CtRootID = 34) "			
			ElseIf Request("CtRootId") = "2" Then			
				sql = sql & "AND (CuDtGeneric.siteId = " & pkStr(Request("CtRootId"), "") & ") AND (CatTreeRoot.vGroup = 'XX')  "								
			Else
				sql = sql & "AND (CuDtGeneric.iCtUnit = " & Request("CtRootId") & ") "								
			End If
		End If
	
		If Request("CtNodeName") <> "" Then
			sql = sql & "AND (CatTreeNode.CatName LIKE " & pkDate("%" & Request("CtNodeName") & "%", "") & ") "
		End If	
		
		If Request("sTitle") <> "" Then
			sql = sql & "AND (CuDTGeneric.sTitle LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "
		End If
		
		If StartDate <> "" And EndDate <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate BETWEEN " & pkStr(StartDate, "") & " AND " & pkStr(EndDate, "") & ") "
		ElseIf StartDate <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate >= " & pkStr(StartDate, "") & ") "
		ElseIf EndDate <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate <= " & pkStr(EndDate, "") & ") "
		End If
		
		If Request("Status") = "Y" Then
			sql = sql & "AND (CuDTGeneric.fCTUPublic = 'Y') "
		Else
			sql = sql & "AND (CuDTGeneric.fCTUPublic <> 'Y') "
		End If
		
		sql = sql & " ORDER BY CuDtGeneric.dEditDate DESC "
				
		'If Request("DeptId") <> "" Then
		'	sql = sql & "AND (CuDTGeneric.iDept = " & pkStr(Request("DeptId"), "") & ")"
		'End If
		'response.write sql
	 session("jigsql")=sql	
	End If
	else
	sql=session("jigsql")
	end if 
	'response.write sql
	
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	
	If Request("CtRootId") = "4" Then		

'----------HyWeb GIP DB CONNECTION PATCH----------
'		Rs.Open sql,KMConn,3,1
Set Rs = KMConn.execute(sql)

'----------HyWeb GIP DB CONNECTION PATCH----------

	Else

'----------HyWeb GIP DB CONNECTION PATCH----------
'		Rs.Open sql,Conn,3,1
Set Rs = Conn.execute(sql)

'----------HyWeb GIP DB CONNECTION PATCH----------

	End If
	PageNumber = Request.QueryString("PageNumber")  '現在頁數

	curRecCount = 0

	If Not Rs.Eof Then 
  	
  	totRec = Rs.Recordcount       '總筆數
   
   	If totRec > 0 Then 
      
      PageSize = cint(Request.QueryString("PageSize"))
      
      If PageSize <= 0 Then  
      	PageSize = 10  
      End If 
      
      Rs.PageSize = PageSize       '每頁筆數

      If cint(PageNumber) < 1 Then 
         PageNumber = 1
      ElseIf cint(PageNumber) > Rs.PageCount Then 
         PageNumber = Rs.PageCount 
      End If            	

      Rs.AbsolutePage = PageNumber
      totPage = Rs.PageCount       '總頁數      
      
      curRecCount = totRec - (PageSize * (PageNumber-1) )    
  		If curRecCount > PageSize Then
  			curRecCount = PageSize
  		End If
      
   	End If    
	End If   

	Function CheckExist(articleid )
	
		sql = "SELECT * FROM [mGIPcoanew].[dbo].[KnowledgeJigsaw] where  Status='Y' and [ArticleId]='"&articleid&"' and parentIcuitem='"&request("gicuitem")&"'" 
		'sql="select * FROM   CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem where   (CuDTGeneric.fCTUPublic = 'Y') and [ArticleId]='"&articleid&"' and parentIcuitem='"&request("gicuitem")&"'" 
		
	  
		Set RsC = Conn.Execute(sql)
		If Not RsC.Eof Then
			CheckExist = true
		Else
			CheckExist = false
		End If
		RsC = Empty		
	End Function
	
function GetPath( sid, id )
	path = ""
	if sid = "3" then					
		topcat = ""
		path = "/knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType={1}&CategoryId={2}"
		sql = "SELECT topCat FROM CuDTGeneric WHERE iCUItem = " & id
		set prs = conn.execute(sql)
		if not prs.eof then
			topcat = prs("topCat")
		end if
		prs.close
		set prs = nothing
		path = replace(path, "{0}", id)
		path = replace(path, "{1}", "A")
		path = replace(path, "{2}", topcat)
	elseif sid = "2" then
		nodeid = ""
		mp = ""
		path = "/subject/ct.asp?xItem={0}&ctNode={1}&mp={2}"
		sql = "SELECT CatTreeNode.CtNodeID, CatTreeNode.CtRootID FROM CatTreeNode "
		sql = sql & "INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
		sql = sql & "INNER JOIN CuDTGeneric ON CtUnit.CtUnitID = CuDTGeneric.iCTUnit WHERE CuDTGeneric.icuitem = " & id
		set prs = conn.execute(sql)
		if not prs.eof then
			nodeid = prs("CtNodeID")
			mp = prs("CtRootID")
		end if
		prs.close
		set prs = nothing
		path = replace(path, "{0}", id)
		path = replace(path, "{1}", nodeid)
		path = replace(path, "{2}", mp)
	elseif sid = "1" then
		nodeid = ""
		path = "/ct.asp?xItem={0}&ctNode={1}&mp=1"
		sql = "SELECT CatTreeNode.CtNodeID FROM CatTreeNode INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
		sql = sql & "INNER JOIN CuDTGeneric ON CtUnit.CtUnitID = CuDTGeneric.iCTUnit WHERE CuDTGeneric.icuitem = " & id
		set prs = conn.execute(sql)
		if not prs.eof then
			nodeid = prs("CtNodeID")			
		end if
		prs.close
		set prs = nothing
		path = replace(path, "{0}", id)
		path = replace(path, "{1}", nodeid)		
	else
		path = "/ct.asp?xItem={0}&ctNode={1}&mp=1"
		path = replace(path, "{0}", id)
		path = replace(path, "{1}", sid)	
	end if
	GetPath = path
end function

FUNCTION pkStr (s, endchar)
  if s="" then
	pkStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkStr="N'" & s & "'" & endchar
  end if
END FUNCTION

FUNCTION pkDate (s, endchar)
  if s="" then
	pkDate = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkDate="'" & s & "'" & endchar
  end if
END FUNCTION
	
	sql2="SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("gicuitem")&"'"
  Set rs2 = conn.Execute(sql2)
%>


<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link type="text/css" rel="stylesheet" href="../css/list.css">
<link type="text/css" rel="stylesheet" href="../css/layout.css">
<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
<script src="../Scripts/AC_ActiveX.js" type="text/javascript"></script>
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
	    <td width="50%" align="left" nowrap class="FormName">農業推薦單元知識拼圖內容管理&nbsp;
		<font size=2>【內容條例清單--<%=tilte%>主題專區 <%=rs2(0)%>】</font>
		</td>
		<td width="50%" class="FormLink" align="right">
			<A href="subject_query.asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&latest=<%=request("latest")%>" title="新增">設定單元內容文章</A>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td align=center colspan=2 width=80% height=230 valign=top>
     <Form name="reg" method="POST" action="subject_setList02.asp?CtRootId=<%=Request("CtRootId")%>&gicuitem=<%=request("gicuitem")%>">
	 <input type="hidden" name="flag" 	value="0" />  
  <p align="center">  
     <% if Request("CtRootId")<>"" then %>
	 <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000"><%=PageNumber%>/<%=totPage%></font>頁|                      
        共<font color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
        <select id="PageNumber" number="PageNumber" size="1" style="color:#FF0000" onchange="PageNumberOnChange()">
                  <% For i = 1 To totPage %>
  							<option value="<%=i%>" <% If i = cint(PageNumber) Then %> selected <% End If %> ><%=i%></option>          
  						<% Next %>        
         </select>      
         頁
		         <% If cint(PageNumber) <> 1 Then %>             
            	|<a href="javascript:PreviousPage()">上一頁</a> 
       			<% End If %>      
       
        		<% If cint(PageNumber) <> Rs.PageCount And cint(PageNumber) < Rs.PageCount Then %> 
            	|<a href="javascript:NextPage()">下一頁</a> 
        		<% End If %>   
		 
		 </font>           
        | 每頁筆數:
		
       <select id="PageSize" name="PageSize" size="1" style="color:#FF0000" onchange="PageSizeOnChange()">       
            <option value="10" <% If PageSize = 10 Then %> selected <% End If %> >10</option>                       
             	<option value="20" <% If PageSize = 20 Then %> selected <% End If %> >20</option>
             	<option value="30" <% If PageSize = 30 Then %> selected <% End If %> >30</option>
             	<option value="50" <% If PageSize = 50 Then %> selected <% End If %> >50</option>
        </select> 
     <%else%>		
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000">0/0</font>頁|                      
        共<font color="#FF0000">0</font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
        <select id="PageNumber" number="PageNumber" size="1" style="color:#FF0000" onchange="PageNumberOnChange()">
                     <option value="10"  selected  >0</option> 
         </select>      
         頁
		        
		 
		 </font>           
        | 每頁筆數:
		
       <select id="PageSize" name="PageSize" size="1" style="color:#FF0000" onchange="PageSizeOnChange()">       
                <option value="10"  selected  >10</option>                       
             	<option value="20" >20</option>
             	<option value="30" >30</option>
             	<option value="50" >50</option>
        </select> 
	 <%end if%>
	 </font>     
  </p>
      <CENTER>
            <TABLE cellSpacing="1" cellPadding="0" width="100%" align="center" border="0">
              <TBODY>
                <TR>
                  <TD vAlign="top" width="95%" colSpan="2" height="230">
                    <CENTER>
                        <TABLE width="100%" id="ListTable">
                          <TBODY>
                            <TR align="left">
                              <th align="middle" width="7%"><div align="center">
                               <div align="center">
							   <input name="selectall" type="button" class="cbutton" onclick="CKBoxSelectAll()" value="全選/全不選">
							   </div>
                              
							  
							  </th>
                              <th>預覽</th>
                              <th>資料庫來源</th>
                              <th>節點名稱</th>
                              <th>資料標題</th>
                              <th>帳號</th>
															<th>單位</th>
                              <th>編修日期</th>
                            </TR>
			     	
          	<% if Request("CtRootId")<>"" then %>
 			<% If totRec > 0 Then %>            	
          	  <% 
			  i = 0	  
			    For i = 1 To PageSize
               'while not Rs.eof
               '   i = i + 1
	          	'if i > (PageNumber - 1) * PageSize and i <= PageNumber * PageSize Then 
			%>
            	<TR>
            		<TD align="middle">
                      <div align="center">                			
            				<% If Request("CtRootId") = "4" Then %>
            					<% If InStr(session("jigcheck"),Rs("REPORT_ID")&"-"&Request("CtRootId")&"-"&Rs("CATEGORY_ID")) >0 or CheckExist(Rs("REPORT_ID")) Then %>
            						<INPUT type="checkbox" name="ckbox" id="ckbox" value="<%=Rs("REPORT_ID")%>-<%=Request("CtRootId")%>-<%=Rs("CATEGORY_ID")%>" checked >
            					<% Else %>
            						<INPUT type="checkbox" name="ckbox" id="ckbox" value="<%=Rs("REPORT_ID")%>-<%=Request("CtRootId")%>-<%=Rs("CATEGORY_ID")%>">
            					<% End If %>
            				<% Else %>
            					<% If InStr(session("jigcheck"),Rs("iCUItem")&"-"&Request("CtRootId")&"-"&Request ("CtRootId")) >0 or CheckExist(Rs("iCUItem")) Then %>
            						<INPUT type="checkbox" name="ckbox" id="ckbox" value="<%=Rs("iCUItem")%>-<%=Request("CtRootId")%>-<%=Request("CtRootId")%>" checked>
            					<% Else %>
            						<INPUT type="checkbox" name="ckbox" id="ckbox" value="<%=Rs("iCUItem")%>-<%=Request("CtRootId")%>-<%=Request("CtRootId")%>">
            					<% End If%>
            				<% End If %>        	
            			</div>            		
								</TD>
								<%
									If Request("CtRootId") = "4" Then
										viewurl = GetPath(Request("CtRootId"), Rs("REPORT_ID"))
									Else
										viewurl = GetPath(Request("CtRootId"), Rs("iCUItem"))
									End If								
								%>
								<TD><A href="<%=session("myWWWSiteURL") & viewurl%>" target="_wMof">View</A></TD>
              	<% If Request("CtRootId") = "4" Then %>
              		<TD>知識庫</TD>
              		<TD><%=Rs("CATEGORY_NAME")%></TD>
              		<TD><%=Rs("SUBJECT")%></TD>
              		<TD><%=Rs("ACTOR_DETAIL_NAME")%>&nbsp;</TD>
              		<TD><%=Rs("PUBLISHER")%>&nbsp;</TD>
              		<TD><%=Rs("ONLINE_DATE")%></TD>
              	<% ElseIf Request("CtRootId") = "3" Then %>	
              		<TD>知識家</TD>
              		<TD><%=Rs("CatName")%></TD>
              		<TD><%=Rs("sTitle")%></TD>
              		<TD><%=Rs("realname")%>&nbsp;</TD>
              		<TD>知識家</TD>
              		<TD><%=Rs("dEditDate")%></TD>
              	<% Else %>
              		<% If Request("CtRootId") = "1" Then %>
              			<TD>入口網</TD>
              		<% ElseIf Request("CtRootId") = "2" Then %>
              			<TD>主題網</TD>
              		<% End If %>
              		<TD><%=Rs("CatName")%></TD>
              		<TD><%=Rs("sTitle")%></TD>
              		<TD><%=Rs("UserName")%></TD>
              		<TD><%=Rs("deptName")%></TD>
              		<TD><%=Rs("dEditDate")%></TD>
              	<% End If %>
            	</TR>
            	<%
				Rs.MoveNext
	         				If Rs.Eof Then Exit For 
				next
        	 				'end if
		              'Rs.movenext
	                   '   wend
  						%>
            <% Else %>
            	<TR><TD align="middle" colspan="8"><div align="center">此條件下無資料</div></TD></TR>                           
            <% End If %> 
            <% end if%>			
            </TBODY>
                        </TABLE>
                        
                    </CENTER>
            </TABLE>
        </CENTER>
      </form>

 
    <input name="button" type="button" class="cbutton" onClick="SaveItems()" value="編修存檔">
    <input name="button" type="button" class="cbutton" onClick="location.href='subjectPubList.asp?iCUItem=<%=Request("iCUItem") %>'" value="回上層">
    <INPUT type="button" class="cbutton" onClick="ResetForm()" value="重設" )?></td>
  </tr>
</table>
</body>
</html>


<script language="javascript">
	
	function PageNumberOnChange() 
	{
	
		var doc = document.reg;
		var flag = false;
		var checkitems = "";
		var uncheckitems = "";
		for( var i = 0; i < <%=curRecCount%>; i++ ) {
				if( doc.ckbox[i].checked ) {
					checkitems += doc.ckbox[i].value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox[i].value + ";";	
				}
			}
			
			
			window.location.href="subject_setList01.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&PageNumber=" + doc.PageNumber.options[doc.PageNumber.selectedIndex].value + "&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&CtRootId="+<%=Request("CtRootId")%>+"&check=" + checkitems + "&uncheck=" + uncheckitems ;
		//doc.action = "subject_setList.asp?iCUItem=<%=Request("iCUItem")%>&PageNumber=" + doc.PageNumber.options[doc.PageNumber.selectedIndex].value + "&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&CtRootId="+<%=Request("CtRootId")%>;
		//doc.submit();
	}
	function PageSizeOnChange() 
	{
		var doc = document.reg;
		doc.action = "subject_setList.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&PageNumber=<%=PageNumber%>&PageSize=" + doc.PageSize.options[doc.PageSize.selectedIndex].value+"&CtRootId="+<%=Request("CtRootId")%>;
		doc.submit();
	}
	function PreviousPage()
	{
		var doc = document.reg;
		var flag = false;
		var checkitems = "";
		var uncheckitems = "";
		for( var i = 0; i < <%=curRecCount%>; i++ ) {
				if( doc.ckbox[i].checked ) {
					checkitems += doc.ckbox[i].value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox[i].value + ";";	
				}
			}
			
			
			window.location.href="subject_setList01.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&PageNumber=<%=PageNumber-1%>&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&CtRootId="+<%=Request("CtRootId")%>+"&check=" + checkitems + "&uncheck=" + uncheckitems ;
	}
	function NextPage()
	{
		var doc = document.reg;
		var flag = false;
		var checkitems = "";
		var uncheckitems = "";
		for( var i = 0; i < <%=curRecCount%>; i++ ) {
				if( doc.ckbox[i].checked ) {
					checkitems += doc.ckbox[i].value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox[i].value + ";";	
				}
			}
			
			
			window.location.href="subject_setList01.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&PageNumber=<%=PageNumber+1%>&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&CtRootId="+<%=Request("CtRootId")%>+"&check=" + checkitems + "&uncheck=" + uncheckitems ;
	}	
	function CKBoxSelectAll()
	{
		var doc = document.reg;
		var flag = false;
		if( doc.flag.value == "0" ) {
			for( var i = 0; i < <%=curRecCount%>; i++ ) {			
				doc.ckbox[i].checked = true;			
			}
			doc.flag.value = "1";
		}
		else {
			for( var i = 0; i < <%=curRecCount%>; i++ ) {			
				doc.ckbox[i].checked = false;			
			}
			doc.flag.value = "0";
		}
	}
	function ResetForm()
	{
	window.location.href= "subject_setList.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&PageNumber=1&PageSize=10";
    }
		
	function SaveItems()
	{
		var doc = document.reg;
		var flag = false;
		var checkitems = "";
		var uncheckitems = "";
		var count = <%=curRecCount%>;
		if (count == 1 ) {
			if( doc.ckbox.checked ) {
				checkitems += doc.ckbox.value + ";";	
			}		
		}
		else {
			for( var i = 0; i < count; i++ ) {
				if( doc.ckbox[i].checked ) {
					checkitems += doc.ckbox[i].value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox[i].value + ";";	
				}
			}
		}	
			
			window.location.href="subject_setList02.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&CtRootId=<%=Request("CtRootId")%>&check=" + checkitems + "&uncheck=" + uncheckitems ;
		
	}
	
</script>
