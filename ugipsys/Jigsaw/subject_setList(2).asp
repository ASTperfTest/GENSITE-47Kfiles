<%@  codepage="65001" %>
<% Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<%
response.write session("jigcheck1")
iCUItem = request("iCUItem")
gicuitem = request("gicuitem")
Page = request("Page")
sql = "SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("iCUItem")&"'"
Set rs = conn.Execute(sql)
sql1 = "SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("gicuitem")&"'"
Set rs1 = conn.Execute(sql1)
sql2 = "SELECT CuDTGeneric.sTitle, KnowledgeJigsaw.gicuitem as gicuitem, KnowledgeJigsaw.orderArticle , "
sql2 = sql2 & "CuDTGeneric.xBody,CuDTGeneric.iEditor,convert(varchar, CuDTGeneric.xpostdate, 111) as xpostdate,  "
sql2 = sql2 & "KnowledgeJigsaw.CtRootId, KnowledgeJigsaw.CtUnitId as CtUnitId, KnowledgeJigsaw.ArticleId as ArticleId, "
sql2 = sql2 & "KnowledgeJigsaw.Status, KnowledgeJigsaw.parentIcuitem, CuDTGeneric.dEditDate, KnowledgeJigsaw.path "
sql2 = sql2 & "FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem WHERE (KnowledgeJigsaw.Status = 'y') and ( KnowledgeJigsaw.parentIcuitem='"&gicuitem&"') ORDER BY KnowledgeJigsaw.orderArticle DESC,dEditDate,xpostdate ASC"
set rs2=conn.Execute(sql2)
'response.write iCUItem & ";" & gicuitem

Function CheckExist(articleid,CtRootId,CtUnitId)

	if CtRootId = 4 then 
		sql4 = "SELECT CATEGORY_NAME as CATEGORY_NAME  FROM [coa].[dbo].[CATEGORY] where [DATA_BASE_ID]='DB020' and [CATEGORY_ID]= '"&CtUnitId&"'"
		Set RsC1 = KMConn.Execute(sql4)
		if not RsC1.eof then 
			CheckExist=RsC1("CATEGORY_NAME")
		else
			CheckExist=""
		end if 
	else
		sql3 = "SELECT CtUnit.CtUnitName as CtUnitName FROM  CuDTGeneric INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID WHERE (CuDTGeneric.iCUItem = '"&articleid&"')"
		Set RsC = Conn.Execute(sql3)
		if not RsC.eof then 
			CheckExist=RsC("CtUnitName")
    else
			CheckExist=""
		end if 
	end if  
		
End Function

	page = 1
	pagecount1 = request("PageSize")
	if pagecount1 = "" then
		pagecount = 10
	else
		pagecount=pagecount1
	end if

	sql3="SELECT count(*) FROM [mGIPcoanew].[dbo].[KnowledgeJigsaw]where Status ='Y' and [parentIcuitem]='"&gicuitem&"'"
	set rs5 = conn.execute(sql3)

	if int(rs5(0)) <= int(pagecount) then 
		page = 1
	else
		if request("Page")="" then
			page = 1
		else
			Page= request("Page")
		end if 
	end if 

	totalpage = 1
	if rs5(0) > 0 then
		totalpage = rs5(0) \ pagecount
		'response.write(totalpage)	
		if rs5(0) mod pagecount <> 0 then	totalpage = totalpage + 1
	end if
  curRecCount = 0
  curRecCount = rs5(0) - (pagecount * (Page-1) )    
	'response.write curRecCount
	If int(curRecCount) > int(pagecount) Then
  	curRecCount = pagecount
  	'response.write curRecCount
	End If
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <link type="text/css" rel="stylesheet" href="../css/list.css">
    <link type="text/css" rel="stylesheet" href="../css/layout.css">
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css">

    <script src="../Scripts/AC_ActiveX.js" type="text/javascript"></script>

    <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>

</head>
<body>
<form method="POST" name="reg" action="subject_setList(2)01.asp?gicuitem=<%=request("gicuitem")%>&iCUItem=<%=request("iCUItem")%>">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="50%" align="left" nowrap class="FormName">農業推薦單元知識拼圖內容管理&nbsp; <font size="2">【內容條例清單--<%=rs("stitle")%>主題專區
                <%=rs1("stitle")%>
                】</font> </td>
            <td width="50%" class="FormLink" align="right">
                <!--A href="subject_query.htm" title="新增">設定單元內容文章</A-->
                <a href="Javascript:window.history.back();" title="回前頁">回前頁</a>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="2"></td>
        </tr>
        <tr>
            <td width="100%" colspan="2">
                <hr noshade size="1" color="#000080">
            </td>
        </tr>
        <tr>
            <td align="center" colspan="2" width="80%" height="230" valign="top">
                <center>
                    <table cellspacing="1" cellpadding="0" width="100%" align="center" border="0">
                        <tbody>
                            <tr>
                                <td valign="top" colspan="2" height="230">
                                    <table border="0" width="100%">
                                        <tr>
                                            <td width="30%">
                                                <%if request("ActionType")="setLink" then %>
                                                <fieldset style="width: 300px"><legend>新增資源推薦聯結</legend>
                                                    <table style="font-size: 12px">
                                                        <tr>
                                                            <td>標題：</td>
                                                            <td colspan="2"><input type="text" name="title" maxlength="30" /></td>
                                                        </tr>
                                                        <tr>
                                                            <td>網址：</td>
                                                            <td><input type="text" name="Url" maxlength="100" /></td>
                                                            <td><input type="button" name="urlsubmit" onclick="SubmitURL()" value="送出"/></td>
                                                        </tr>
                                                    </table>
                                                    <input type="hidden" name="AddLinkAction" id="IdAddLinkAction" value="" />
                                                </fieldset>
                                                <%end if %>
                                            </td>
                                            <td width="70%" align="left">
                                                <p><font size="2" color="rgb(63,142,186)">第 <font size="2" color="#FF0000">
                                                    <%=page%>/<%=totalpage%></font>頁| 共<font color="#FF0000"><%=rs5(0)%></font> <font size="2" color="rgb(63,142,186)">筆| 跳至第
                                                        <select id="select" name="select" size="1" style="color: #FF0000" onchange="PageNumberOnChange()">
                                                            <%
										                    n = 1
										                    while n <= totalpage
											                    response.write "<option value='" & n & "'"
											                    if int(n) = int(page) then	response.write " selected"
											                    response.write ">" & n & "</option>"
											                    n = n + 1
										                    wend        
                                                            %>
                                                        </select>
                                                        頁</font> | 每頁筆數:
                                                    <select id="PageSize" name="PageSize" size="1" style="color: #FF0000" onchange="PageSizeOnChange()">
                                                        <option value="10" <% If request("PageSize") = 10 Then %> selected <% End If %>>10</option>
                                                        <option value="30" <% If request("PageSize") = 30 Then %> selected <% End If %>>30</option>
                                                        <option value="50" <% If request("PageSize")= 50 Then %> selected <% End If %>>50</option>
                                                    </select>
                                                </font></p>
                                            </td>
                                        </tr>
                                    </table>
                                    <center>
                                        
                                        <input type="hidden" name="flag" value="0" />
                                        <input type="hidden" name="submitTask2" value="">
                                        <input type="hidden" name="CalendarTarget">
                                        <table width="100%" id="ListTable">
                                            <tbody>
                                                <tr align="left">
                                                    <th align="middle" width="7%">
                                                        <div align="center">
                                                            <input name="selectall" type="button" class="cbutton" onclick="CKBoxSelectAll()" value="全選/全不選"></div>
                                                    </th>
                                                    <%if request("ActionType")="setLink" then %>
                                                    <th>預覽</th>
                                                    <th>標題</th>
                                                    <th>URL</th>
                                                    <th>順序</th>
                                                    <th>編修日期</th>
													<%elseif request("ActionType")="ManageDiscussion" then%>
													<th>留言人</th>
                                                    <th>留言日期</th>
                                                    <th>留言內容</th>
                                                    
                                                    <th>編修日期</th>
                                                    <%else %>
                                                    <th>預覽</th>
                                                    <th>資料庫來源</th>
                                                    <th>單元名稱</th>
                                                    <th>資料標題</th>
                                                    <th>順序</th>
                                                    <th>編修日期</th>
                                                    <%end if%>
                                                </tr>
                                                <%
										          i = 0	  
			                                      while not rs2.eof
		                                          i = i + 1
		                                          if i > (page - 1) * pagecount and i <= page * pagecount Then
                                                %>
                                                <%if request("ActionType")="setLink" then %>
                                                <tr>
                                                    <td align="middle">
                                                        <div align="center">
                                                            <% If InStr(session("jigcheck1"),rs2("gicuitem")) Then %>
                                                            <input type="checkbox" name="ckbox" id="ckbox" value="<%=rs2("gicuitem")%>" checked>
                                                            <% Else %>
                                                            <input type="checkbox" name="ckbox" id="ckbox" value="<%=rs2("gicuitem")%>">
                                                            <% End If %>
                                                        </div>
                                                    </td>
                                                    <td><a href="<%=rs2("path")%>" target="_blank">View</a></td>
                                                    <td><%=rs2("sTitle")%></td>
                                                    <td><%=rs2("path")%></td>
                                                    <td>
                                                        <span class="eTableContent">
                                                        <input type="text" name="<%=rs2("gicuitem")%>" value="<%=rs2("orderArticle")%>" size="5"></span>
                                                    </td>
                                                    <td>
                                                        <%=rs2("dEditDate")%></td>
                                                </tr>
												<%elseif request("ActionType")="ManageDiscussion" then%>
												<tr>
                                                    <td align="middle">
                                                        <div align="center">
                                                            <% If InStr(session("jigcheck1"),rs2("gicuitem")) Then %>
                                                            <input type="checkbox" name="ckbox" id="ckbox" value="<%=rs2("gicuitem")%>" checked>
                                                            <% Else %>
                                                            <input type="checkbox" name="ckbox" id="ckbox" value="<%=rs2("gicuitem")%>">
                                                            <% End If %>
                                                        </div>
                                                    </td>
                                                    <td><%=rs2("iEditor")%></td>
                                                    <td><%=rs2("xpostdate")%></td>
                                                    <td><%=rs2("xBody")%></td>                                                    
                                                    <td>
                                                        <%=rs2("dEditDate")%></td>
                                                </tr>
                                                <%else%>
                                                <tr>
                                                    <td align="middle">
                                                        <div align="center">
                                                            <% If InStr(session("jigcheck1"),rs2("gicuitem")) Then %>
                                                            <input type="checkbox" name="ckbox" id="ckbox" value="<%=rs2("gicuitem")%>" checked>
                                                            <% Else %>
                                                            <input type="checkbox" name="ckbox" id="ckbox" value="<%=rs2("gicuitem")%>">
                                                            <% End If %>
                                                        </div>
                                                    </td>
                                                    <td><a href="<%=session("myWWWSiteURL") & rs2("path")%>" target="_wMof">View</a></td>
                                                    <td>
                                                        <%
											            select case(rs2("CtRootId"))
												            case 1
													            response.write "入口網"
												            case 2
													            response.write "主題館"
												            case 3
													            response.write "知識家"
												            case 4
													            response.write "知識庫"
												            case else
													            response.write "&nbsp;"
											            end select
                                                        %>
                                                    </td>
                                                    <td>
                                                        <%  
											                response.write CheckExist(rs2("ArticleId"),rs2("CtRootId"),rs2("CtUnitId"))
                                                        %>
                                                    </td>
                                                    <td>
                                                        <%=rs2("sTitle")%></td>
                                                    <td>
                                                        <span class="eTableContent">
                                                        <input type="text" name="<%=rs2("gicuitem")%>" value="<%=rs2("orderArticle")%>" size="5"></span>
                                                    </td>
                                                    <td>
                                                        <%=rs2("dEditDate")%></td>
                                                </tr>
                                                <%end if %>                                                
                                                <%
                                                end if
		                                            rs2.movenext
	                                            wend
                                                %>
                                            </tbody>
                                        </table>
                                        <input name="button" type="button" class="cbutton" onclick="SaveItems()" value="刪除選擇">
										<%if request("ActionType")<>"ManageDiscussion" then %>
                                        <input name="button" type="submit" class="cbutton" onclick="location.href='subject_setList(2)01.asp?gicuitem=<%=request("gicuitem")%>&iCUItem=<%=request("iCUItem")%>&ActionType=<%=request("ActionType") %>'" value="編修存檔">
                                        <%end if %> 
                                    </center>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </center>
        </tr>
    </table>
    </form>
</body>
</html>

<script language="JavaScript">
function PageNumberOnChange(){
var objSelect =document.reg.select;
	var doc = document.reg;
   window.location.href="subject_setList(2).asp?ActionType=<%=request("ActionType") %>&iCUItem=<%=request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&page="+objSelect.options[objSelect.selectedIndex].value+"&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value;
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
function PageSizeOnChange() 
	{
		var doc = document.reg;
		var objSelect =document.reg.select;
		doc.action = "subject_setList(2).asp?ActionType=<%=request("ActionType") %>&iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&Page="+objSelect.options[objSelect.selectedIndex].value+"&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value;
		doc.submit();
	}
	function PageNumberOnChange() 
	{
    	
	    var objSelect =document.reg.select;
	    var doc = document.reg;
 	
		var flag = false;
		var checkitems = "";
		var uncheckitems = "";
		
		if (<%=curRecCount%> > 1)
		{
		    for( var i = 0; i < <%=curRecCount%>; i++ ) {
				if( doc.ckbox[i].checked ) {
					checkitems += doc.ckbox[i].value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox[i].value + ";";	
				}
			}
		}
		else
		{
				if( doc.ckbox.checked ) {
					checkitems += doc.ckbox.value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox.value + ";";	
				}		
		}
		
	
	        window.location.href="subject_setList(2).asp?ActionType=<%=request("ActionType") %>&iCUItem=<%=request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&page="+objSelect.options[objSelect.selectedIndex].value+"&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&check=" + checkitems + "&uncheck=" + uncheckitems ;
			//mark by bob
			//window.location.href="subject_setList(2)03.asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&page="+objSelect.options[objSelect.selectedIndex].value+"&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&check=" + checkitems + "&uncheck=" + uncheckitems ;
			
			
			//window.location.href="subject_setList(2).asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&page="+objSelect.options[objSelect.selectedIndex].value+"&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value;
			//window.location.href="subject_setList01.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&PageNumber=" + doc.PageNumber.options[doc.PageNumber.selectedIndex].value + "&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&CtRootId="+<%=Request("CtRootId")%>+"&check=" + checkitems + "&uncheck=" + uncheckitems ;
		//doc.action = "subject_setList.asp?iCUItem=<%=Request("iCUItem")%>&PageNumber=" + doc.PageNumber.options[doc.PageNumber.selectedIndex].value + "&PageSize="+ doc.PageSize.options[doc.PageSize.selectedIndex].value+"&CtRootId="+<%=Request("CtRootId")%>;
		//doc.submit();
	}
	function SubmitURL()
	{
	    reg.AddLinkAction.value="AddLink";
	    if (reg.title.value=='') 
	    {
	        alert('標題不可為空白');
	        reg.title.focus();
	        return;
	    }
	    if (reg.Url.value=='') 
	    {
	        alert('網址不可為空白');
	        reg.Url.focus();
	        return;
	    }	    
	    reg.submit();
	}
	
	function SaveItems()
	{
		var objSelect =document.reg.select;
	
		var doc = document.reg;
		var flag = false;
		var checkitems = "";
		var uncheckitems = "";

		if (<%=curRecCount%> > 1)
		{
		    for( var i = 0; i < <%=curRecCount%>; i++ ) {
				if( doc.ckbox[i].checked ) {
					checkitems += doc.ckbox[i].value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox[i].value + ";";	
				}
			}
		}
		else
		{
				if( doc.ckbox.checked ) {
					checkitems += doc.ckbox.value + ";";	
				}			
				else{
					uncheckitems += doc.ckbox.value + ";";	
				}		
		}
			
			
			window.location.href="subject_setList(2)02.asp?iCUItem=<%=Request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>&check=" + checkitems+ "&uncheck=" + uncheckitems ;  ;
		
	}
</script>

