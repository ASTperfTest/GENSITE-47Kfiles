<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="電子報發行清單"
HTProgCode="GW1M51"
HTProgPrefix="ePub"
%>
<html>
	<head>
		<META http-equiv="Content-Type" content="text/html; charset=utf-8">
		<link type="text/css" rel="stylesheet" href="/css/list.css">
		<link type="text/css" rel="stylesheet" href="/css/layout.css">
		<link type="text/css" rel="stylesheet" href="/css/setstyle.css">
	</head>
<!--#INCLUDE FILE="ePubListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%	
	Dim epubid
	epubid = Request("epubId")
	If epubid = "" Then response.write "<script>alert('缺少電子報代碼值');history.back();</script>"				
	
	Set KMConn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	KMConn.Open session("KMODBC")
'Set KMConn = Server.CreateObject("HyWebDB3.dbExecute")
KMConn.ConnectionString = session("KMODBC")
KMConn.ConnectionTimeout=0
KMConn.CursorLocation = 3
KMConn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	
	Dim epubtitle
	sql = "SELECT * FROM EpPub WHERE ePubID = " & epubid
	Set Rs = Conn.Execute(sql)
	If Not Rs.EOF Then
		epubtitle = rs("title")
	End If
	Rs = Empty
	
	sql = ""
	sql = sql & " SELECT * FROM EpPubArticle WHERE (EpPubArticle.EPubId = " & Request("epubid") & ")"
	
	Set Rs = Conn.Execute(sql)

	Dim RecordCount
	Function GetTableContent( myrs )
		
		str = ""
		If myrs("CtRootId") = "1" Or myrs("CtRootId") = "2" Then
			
			sql = ""	
			sql = sql & "SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, "
			sql = sql & "CtUnit.CtUnitName, InfoUser.UserName, CuDTGeneric.dEditDate, Dept.deptName FROM CuDTGeneric INNER JOIN "
			sql = sql & "CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON "
			sql = sql & "CatTreeNode.CtRootID = CatTreeRoot.CtRootID INNER JOIN CtUnit ON "
			sql = sql & "CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN InfoUser ON "
			sql = sql & "CuDTGeneric.iEditor = InfoUser.UserID INNER JOIN Dept ON CuDTGeneric.iDept = Dept.deptID "
			sql = sql & "WHERE (CatTreeRoot.inUse = 'Y') AND  (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') "
			sql = sql & "AND (CatTreeRoot.inUse = 'Y') AND (Dept.inUse = 'Y') "
			sql = sql & "AND (CuDTGeneric.iCuItem = " & myrs("ArticleId") & ")"
			
			set newrs = Conn.Execute(sql)			
			
			if not newrs.eof then
				str = str & "<TD align=""middle""><div align=""center"">"
				str = str & "<input type=""checkbox"" name=""ckbox"" value="""& newrs("iCUItem") & "-" & myrs("CtRootId") & "-" & myrs("CtRootId") & """ checked>"
				str = str & "</div></TD> "
	      str = str & "<TD><A href=""#"" target=""_wMof"">View</A></TD>"
	      If myrs("CtRootId") = "1" Then
	      	str = str & "<TD>入口網</TD>"
	      Else
	      	str = str & "<TD>主題館</TD>"
	    	End If      
	     	str = str & "<TD>" & newrs("CtUnitName") & "</TD>"
	      str = str & "<TD>" & newrs("sTitle") & "</TD>"
	      str = str & "<TD>" & newrs("UserName") & "</TD>"
	      str = str & "<TD>" & newrs("deptName") & "</TD>"
	      str = str & "<TD>" & FormatDateTime(newrs("dEditDate"),2) & "</TD>"
      end if
			newrs.close
      set newrs = nothing
      
    ElseIf myrs("CtRootId") = "3" Then  
      
      sql = ""
			sql = sql & " SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, "
			sql = sql & " CtUnit.CtUnitName, CuDTGeneric.dEditDate, Member.realname FROM CuDTGeneric INNER JOIN CatTreeNode ON "
			sql = sql & " CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON CatTreeNode.CtRootID = CatTreeRoot.CtRootID "
			sql = sql & " INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
			sql = sql & " WHERE (CuDTGeneric.siteId = N'3') AND (CatTreeRoot.inUse = 'Y') AND (CuDTGeneric.fCTUPublic = 'Y') "
			sql = sql & " AND (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') AND (CatTreeRoot.inUse = 'Y') AND (CatTreeNode.CtUnitID = 932) "
			sql = sql & " AND (CuDTGeneric.iCuItem = " & myrs("ArticleId") & ")"
			
			set newrs = Conn.Execute(sql)			
			
			if not newrs.eof then 
				str = str & "<TD align=""middle""><div align=""center"">"
				str = str & "<input type=""checkbox"" name=""ckbox"" value="""& newrs("iCUItem") & "-" & myrs("CtRootId") & "-" & myrs("CtRootId") & """ checked>"
				str = str & "</div></TD> "
	      str = str & "<TD><A href=""#"" target=""_wMof"">View</A></TD>"
	      str = str & "<TD>知識家</TD>"     
	     	str = str & "<TD>" & newrs("CtUnitName") & "</TD>"
	      str = str & "<TD>" & newrs("sTitle") & "</TD>"
	      str = str & "<TD>" & newrs("realname") & "</TD>"
	      str = str & "<TD>知識家</TD>"
	      str = str & "<TD>" & FormatDateTime(newrs("dEditDate"),2) & "</TD>"
      end if 
			newrs.close
      set newrs = nothing
      
    ElseIf myrs("CtRootId") = "4" Then
	
			sql = ""
			IF CBool(myrs("ISFORMER")) THEN
				sql = sql & " SELECT DISTINCT REPORT.REPORT_ID, CATEGORY.CATEGORY_NAME, CATEGORY.CATEGORY_ID, REPORT.SUBJECT, REPORT.PUBLISHER, REPORT.ONLINE_DATE, "
				sql = sql & " ACTOR_INFO.ACTOR_DETAIL_NAME FROM REPORT INNER JOIN ACTOR_INFO ON REPORT.CREATE_USER = ACTOR_INFO.ACTOR_INFO_ID "
				sql = sql & " INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY ON "
				sql = sql & " CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID AND CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID "
				sql = sql & " INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID "
				sql = sql & " WHERE (REPORT.STATUS = 'PUB') AND (REPORT.ONLINE_DATE < GETDATE()) AND (CATEGORY.DATA_BASE_ID = 'DB020') "
				sql = sql & " AND (RESOURCE_RIGHT.ACTOR_INFO_ID = '002') "
				sql = sql & " AND (REPORT.REPORT_ID = '" & myrs("ArticleId") & "')"
				
				Set kmrs = KMConn.Execute(sql)
				if not kmrs.eof then
					str = str & "<TD align=""middle""><div align=""center"">"
					str = str & "<input type=""checkbox"" name=""ckbox"" value="""& kmrs("REPORT_ID") & "-" & myrs("CtRootId") & "-" & kmrs("CATEGORY_ID") & """ checked>"
					str = str & "</div></TD> "
					str = str & "<TD><A href=""#"" target=""_wMof"">View</A></TD>"
					str = str & "<TD>知識庫</TD>"
					str = str & "<TD>" & kmrs("CATEGORY_NAME") & "</TD>"
					str = str & "<TD>" & kmrs("SUBJECT") & "</TD>"
					str = str & "<TD>" & kmrs("ACTOR_DETAIL_NAME") & "</TD>"
					str = str & "<TD>" & kmrs("PUBLISHER") & "</TD>"
					str = str & "<TD>" & FormatDateTime(kmrs("ONLINE_DATE"),2) & "</TD>"
				end if
				kmrs.close
				set kmrs = nothing
			ELSE
					detailInfo = GetDetail("DOCUMENT",myrs("articleid"))
					infos = split(detailInfo, "|")
					str = str & "<TD align=""middle""><div align=""center"">"
					str = str & "<input type=""checkbox"" name=""ckbox"" value="""& myrs("articleid") & "-" & myrs("CtRootId") & "-" & myrs("categoryid") & """ checked>"
					str = str & "</div></TD> "
					Url = session("myWWWSiteURL") & "/category/categorycontent.aspx?ReportId=" & myrs("articleid") & "&CategoryId=" & myrs("categoryid") & "&ActorType=002&kpi=0"
					str = str & "<TD><A href=" & Url & " target=""_blank"">View</A></TD>"
					str = str & "<TD>知識庫</TD>"
					str = str & "<TD>" & GetDetail("CATEGORY_NAME",myrs("categoryid")) & "</TD>"
					str = str & "<TD>" & infos(0) & "</TD>"
					str = str & "<TD>" & GetDetail("USER_NAME",infos(1)) & "</TD>"
					str = str & "<TD>" & GetDetail("USER_OU",infos(1)) & "</TD>"
					str = str & "<TD>" & Replace(Mid(infos(2),1,10),"-","/") & "</TD>"
			END IF
    End If       
    
    GetTableContent = str
	End Function
	
	Function CheckExist( epubid, articleid )
	
		sql = "SELECT * FROM EpPubArticle WHERE EpubId = " & epubid & " AND ArticleId = " & articleid
		Set RsC = Conn.Execute(sql)
		If Not RsC.Eof Then
			CheckExist = true
		Else
			CheckExist = false
		End If
		RsC = Empty		
	End Function
	
	Function GetDetail(typeid, key)
		'建立物件
		 dim xmlhttp
		 Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
		 '要檢查的網址
		 
		 URLs = session("mySiteMMOURL") & "/epaper/epaper_querydetail.asp?typeid=" & typeid & "&KEY=" & key

		 IF typeid = "" OR key = "" Then
		 Response.Write "未傳入必要參數"
		 Response.Write.End()
		 End IF
		 'Response.Write xmlhttp
		 xmlhttp.Open "Get", URLs, false
		 xmlhttp.Send
		 
		 IF xmlhttp.status=404 Then
			  Response.Write "找不到頁面"
		 ELSEIF xmlhttp.status=200 Then
				GetDetail = xmlhttp.responsetext
		 Else
			  Response.Write xmlhttp.status
		 End If
	End Function
%>
	<body>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left" nowrap class="FormName">電子報管理&nbsp;<font size=2>【最新農業知識文章設定清單--<%=epubtitle%>】</td>
			<td width="50%" class="FormLink" align="right">			
				<a href="epaper_query.asp?epubid=<%=epubid%>&eptreeid=<%=Request("eptreeid")%>" target="">新增最新農業知識文章</a>		
				<a href="Javascript:window.history.back();" title="回前頁">回前頁</a>	    
			</td>
		</tr>
		<tr><td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td></tr>
  	<tr><td class="Formtext" colspan="2" height="15"></td></tr>
  	<tr>
    	<td align=center colspan=2 width=80% height=230 valign=top>
    		<form method="POST" name="iForm">    
    			<input type="hidden" name="flag" 				value="0" />     		
        	<CENTER>
        	<TABLE cellSpacing="1" cellPadding="0" width="100%" align="center" border="0">
        	<TBODY>
        	<TR>
        		<TD vAlign="top" width="95%" colSpan="2" height="230">
        			<!--div id="Page">
          			<p align="center">
          			<font size="2" color="rgb(63,142,186)">第
     						<font size="2" color="#FF0000"><%=PageNumber%>/<%=totPage%></font>頁|共<font color="#FF0000"><%=totRec%></font>
       					<font size="2" color="rgb(63,142,186)">筆| 跳至第       
  							<select id="PageNumber" number="PageNumber" size="1" style="color:#FF0000" onchange="PageNumberOnChange()">
  							<% For i = 1 To totPage %>
  								<option value="<%=i%>" <% If i = cint(PageNumber) Then %> selected <% End If %> ><%=i%></option>          
  							<% Next %>           		
         				</select>  
         				<% If cint(PageNumber) <> 1 Then %>             
            			|<a href="javascript:PreviousPage()">上一頁</a> 
       					<% End If %>             
        				<% If cint(PageNumber) <> Rs.PageCount And cint(PageNumber) < Rs.PageCount Then %> 
            			|<a href="javascript:NextPage()">下一頁</a> 
        				<% End If %>   
         				頁</font>                       
        				| 每頁筆數:
       					<select id="PageSize" name="PageSize" size="1" style="color:#FF0000" onchange="PageSizeOnChange()">            
             			<option value="10" <% If PageSize = 10 Then %> selected <% End If %> >10</option>                       
             			<option value="20" <% If PageSize = 20 Then %> selected <% End If %> >20</option>
             			<option value="30" <% If PageSize = 30 Then %> selected <% End If %> >30</option>
             			<option value="50" <% If PageSize = 50 Then %> selected <% End If %> >50</option>
        				</select>     
     						</font>
     					</div-->
            	<CENTER>
            	<TABLE width="100%" id="ListTable">
            	<TBODY>
            	<TR align="left">
            		<th align="middle" width="7%">
	            		<div align="center"><input name="selectall" type="button" class="cbutton" onclick="CKBoxSelectAll()" value="全選/全不選"></div>
            		</th>
              	<th>預覽</th>
              	<th>目錄樹</th>
              	<th>單元名稱</th>
              	<th>資料標題</th>
              	<th>帳號</th>
              	<th>單位</th>
              	<th>編修日期</th>
            	</TR>
            	<% If Not Rs.Eof Then %>
            		<% While Not Rs.Eof %>
            			<% RecordCount = RecordCount + 1 %>
            			<TR>
            				<%=GetTableContent(Rs)%>
            			</TR>
            		<%
        	 					Rs.MoveNext	         				
    							Wend 
  							%>
            	<% Else %>
            		<TR><TD align="middle" colspan="8"><div align="center">此條件下無資料</div></TD></TR>                           
            	<% End If %>            
            	</TBODY>
           		</TABLE>
            	</CENTER>
         		</TABLE>
        	</CENTER>
      	</form>
      	<input name="button" type="button" class="cbutton" onClick="SaveItems()" value="編修存檔">
    		<input type="button" class="cbutton" onClick="ResetForm()" value="重設"></td>
      </tr>
			</table>
		</body>
	</html>
	<script language="javascript">
		
	
		function CKBoxSelectAll()
		{
			var doc = document.iForm;
			var flag = false;
			if( doc.flag.value == "0" ) {
				for( var i = 0; i < <%=RecordCount%>; i++ ) {						
					doc.ckbox[i].checked = true;			
				}
				doc.flag.value = "1";
			}
			else {
				for( var i = 0; i < <%=RecordCount%>; i++ ) {			
					doc.ckbox[i].checked = false;			
				}
				doc.flag.value = "0";
			}
		}
		function ResetForm()
		{
			var doc = document.iForm;
			doc.Items.value = "";
			doc.action = "epaper_setList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>";
			doc.submit();
		}
		function SaveItems()
		{
			var doc = document.iForm;
			var checkitems = "";
			var uncheckitems = "";
			
			for( var i = 0; i < <%=RecordCount%>; i++ ) {
				if( doc.ckbox[i].checked ) {
					checkitems += doc.ckbox[i].value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox[i].value + ";";	
				}				
			}		
			AddDelArticle( checkitems, uncheckitems);
		}
		function AddDelArticle( checkitems, uncheckitems )
		{
			var oxmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
			var httpStr = "../epaper/AddArticle.asp?epubid=<%=Request("epubid")%>&check=" + checkitems + "&uncheck=" + uncheckitems;
			oxmlhttp.open( "GET", httpStr, false );
	  	oxmlhttp.send();
	  	var oRtn = oxmlhttp.responseText;  	  		
	  	if(oRtn == 1) {  				
	  		alert("編修成功");
	  		document.iForm.action = "ePubList.asp?epTreeID=<%=Request("eptreeid")%>";
	  		document.iForm.submit();
	  	}  	
		}
	</script>
