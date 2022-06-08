<%@  codepage="65001" %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="news"
HTProgPrefix="NewsApprove" 

kwTitle =request("kwTitle")
subjectId = request("subjectId")
validate  = request("validate")
if validate="" then validate = "P"
sort  = request("sort")
if sort="" then sort = "title"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString = session("ODBCDSN")
Conn.CursorLocation = 3
Conn.open

	Dim pass : pass = request.querystring("pass")
	Dim notpass : notpass = request.querystring("notpass")
	Dim id : id = request.querystring("id")
	Dim memberId : memberId = session("userId")
	
	if pass = "Y" then
		passArticle
	elseif notpass = "Y" then
		notPassArticle
	else	
		Dim subject, subjectArr
		Dim flag : flag = false
		'---找出此user所管理的主題館---
		sql = "SELECT CatTreeRoot.CtRootID, CatTreeRoot.CtRootName " & _
					"FROM NodeInfo INNER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID " & _
					"WHERE (NodeInfo.owner = " & pkstr(memberId, "") & " OR Exists(select UserId from infouser where UserId= " & pkstr(memberId, "") & " and charindex('SysAdm',ugrpID) > 0))"
		set rs = conn.execute(sql)
		while not rs.eof 
			flag = true
			subject = subject & rs("CtRootID") & ","
			rs.movenext
		wend
		if flag then
			showForm
		else 
			showDoneBox "無權限", "false"
		end if
	end if
%>
<%
Sub passArticle
	Dim items : items = left(id, len(id) - 1)
	sql = "UPDATE CuDTGeneric SET fCTUPublic = 'Y', xAbstract=" & pkstr(memberId, "") & " WHERE iCUItem IN (" & items & ")"
	conn.execute(sql)
	showDoneBox "審核通過", "true"
End Sub
%>
<%
Sub notPassArticle
	Dim items : items = left(id, len(id) - 1)
	sql = "DELETE FROM CuDTGeneric WHERE iCUItem IN (" & items & ")"
	conn.execute(sql)
	showDoneBox "審核不通過", "true"
End Sub
%>
<% Sub showDoneBox(lMsg, btype) %>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
    <title>編修表單</title>
</head>
<body>

    <script type="text/javascript">
	  alert("<%=lMsg%>");
		<% if btype = "true" then %>		    		    
			document.location.href="NewsApproveList.asp?nowpage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>&subjectId=<%=subjectId %>&validate=<%=validate %>&sort=<%=sort %>&kwTitle=<%=server.urlencode(kwTitle)%>";
		<% else %>
			history.go(-1)
		<% end if %>
		
	</script>

</body>
</html>
<% End sub %>
<% 
Sub showForm 

	cSql = "SELECT COUNT(*) FROM CuDTGeneric INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID " & _
				 "INNER JOIN CatTreeRoot INNER JOIN NodeInfo ON NodeInfo.CtrootID = CatTreeRoot.CtRootID " & _
				 "INNER JOIN KnowledgeToSubject ON CatTreeRoot.CtRootID = KnowledgeToSubject.subjectId " & _
				 "INNER JOIN CatTreeNode ON KnowledgeToSubject.ctNodeId = CatTreeNode.CtNodeID ON CtUnit.CtUnitID = CatTreeNode.CtUnitID " & _
				 "WHERE (NodeInfo.owner = " & pkstr(memberId, "") & " OR Exists(select UserId from infouser where UserId= " & pkstr(memberId, "") & " and charindex('SysAdm',ugrpID) > 0)) AND (KnowledgeToSubject.status = 'Y') AND (CuDTGeneric.fCTUPublic = 'P')"
'//(NodeInfo.owner = " & pkstr(memberId, "") & ") AND


	fSql = "select CatTreeRoot.CtRootName, CatTreeRoot.pvXdmp, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, KnowledgeToSubject.ctNodeId, " & _
				 "CuDTGeneric.xURL, CuDTGeneric.showType, CtUnit.CtUnitName, CONVERT(varchar,CuDTGeneric.xPostDate,111) AS xPostDate " & _
				 "FROM CuDTGeneric INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID " & _
				 "INNER JOIN CatTreeRoot INNER JOIN NodeInfo ON NodeInfo.CtrootID = CatTreeRoot.CtRootID " & _
				 "INNER JOIN KnowledgeToSubject ON CatTreeRoot.CtRootID = KnowledgeToSubject.subjectId " & _
				 "INNER JOIN CatTreeNode ON KnowledgeToSubject.ctNodeId = CatTreeNode.CtNodeID ON CtUnit.CtUnitID = CatTreeNode.CtUnitID " & _
				 "WHERE (NodeInfo.owner = " & pkstr(memberId, "") & " OR Exists(select UserId from infouser where UserId= " & pkstr(memberId, "") & " and charindex('SysAdm',ugrpID) > 0)) AND (KnowledgeToSubject.status = 'Y') " & _
				 "AND (CuDTGeneric.fCTUPublic = 'P') ORDER BY CuDTGeneric.sTitle, CuDTGeneric.xPostDate,CatTreeRoot.CtRootName DESC "
'//(NodeInfo.owner = " & pkstr(memberId, "") & ") AND		


fSQL =""
fSql = fSql & vbcrlf & "declare @user nvarchar(50)"
fSql = fSql & vbcrlf & "set @user = " & pkstr(memberId, "")
fSql = fSql & vbcrlf & "select "
fSql = fSql & vbcrlf & "		CatTreeRoot.CtRootName"
fSql = fSql & vbcrlf & "	, CatTreeRoot.pvXdmp"
fSql = fSql & vbcrlf & "	, CuDTGeneric.iCUItem"
fSql = fSql & vbcrlf & "	, CuDTGeneric.sTitle"
fSql = fSql & vbcrlf & "	, KnowledgeToSubject.ctNodeId"
fSql = fSql & vbcrlf & "	, isnull((select userName from infouser where UserId=CuDTGeneric.xAbstract),'') as Signer"
fSql = fSql & vbcrlf & "	, CuDTGeneric.xURL"
fSql = fSql & vbcrlf & "	, CuDTGeneric.showType"
fSql = fSql & vbcrlf & "	, CuDTGeneric.fCTUPublic"
fSql = fSql & vbcrlf & "	, CtUnit.CtUnitName"
fSql = fSql & vbcrlf & "	, CONVERT(varchar,CuDTGeneric.xPostDate,111) AS xPostDate "
fSql = fSql & vbcrlf & "FROM CuDTGeneric "
fSql = fSql & vbcrlf & "INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID "
fSql = fSql & vbcrlf & "INNER JOIN CatTreeRoot "
fSql = fSql & vbcrlf & "INNER JOIN NodeInfo ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
fSql = fSql & vbcrlf & "INNER JOIN KnowledgeToSubject ON CatTreeRoot.CtRootID = KnowledgeToSubject.subjectId "
fSql = fSql & vbcrlf & "INNER JOIN CatTreeNode ON KnowledgeToSubject.ctNodeId = CatTreeNode.CtNodeID ON CtUnit.CtUnitID = CatTreeNode.CtUnitID "
fSql = fSql & vbcrlf & "WHERE "
fSql = fSql & vbcrlf & "	(NodeInfo.owner = @user OR Exists(select UserId from infouser where UserId= @user and charindex('SysAdm',ugrpID) > 0)) "
fSql = fSql & vbcrlf & "AND (KnowledgeToSubject.status = 'Y') "
if kwTitle <>"" then
    fSql = fSql & vbcrlf & "AND (CuDTGeneric.sTitle like '%' + " & pkstr(kwTitle, "") & " + '%') "
else
    if validate <> "" then
        fSql = fSql & vbcrlf & "AND (CuDTGeneric.fCTUPublic = '"& validate &"') "
    end if
    if subjectId <> "" then
        fSql = fSql & vbcrlf & "AND (KnowledgeToSubject.subjectId = '" & subjectId & "') "
    end if
end if
select case sort
    case "title"
        fSql = fSql & vbcrlf & "ORDER BY CuDTGeneric.sTitle, CuDTGeneric.xPostDate,CatTreeRoot.CtRootName DESC "
    case "subject"
        fSql = fSql & vbcrlf & "ORDER BY CatTreeRoot.CtRootName DESC ,CuDTGeneric.sTitle, CuDTGeneric.xPostDate"
    case "date"
        fSql = fSql & vbcrlf & "ORDER BY CuDTGeneric.xPostDate, CuDTGeneric.sTitle, CatTreeRoot.CtRootName DESC "
end select
if isnumeric(Request("pagesize")) then
    PerPageSize = cint(Request("pagesize"))
    if PerPageSize <= 0 then  PerPageSize = 15  
else
    PerPageSize = 15  
end if 
 

nowPage = Request("nowPage")  '現在頁數
if not isnumeric(nowPage) then nowPage=1
if nowPage<=0 then nowPage=1		

Set RSreg = Server.CreateObject("ADODB.RecordSet")
'----------HyWeb GIP DB CONNECTION PATCH----------
	
    set RSreg = conn.execute(fsql)
'----------HyWeb GIP DB CONNECTION PATCH----------



    totPage=0
    RSreg.PageSize = PerPageSize
    totRec = RSreg.recordcount
	if Not RSreg.eof then
	    if cint(nowPage)>RSreg.pagecount then nowPage=RSreg.pagecount 	                   
        RSreg.AbsolutePage = nowPage
        totPage = RSreg.pagecount
	end if   
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="/css/list.css" rel="stylesheet" type="text/css">
    <link href="/css/layout.css" rel="stylesheet" type="text/css">
    <title>資料管理／主題館新聞發佈審核</title>
</head>
<body>
    <div id="FuncName">
        <h1>
            資料管理／主題館新聞發佈審核</h1>
        <div id="Nav">
        </div>
        <div id="ClearFloat">
        </div>
    </div>
    <form id="iForm" name="iForm" method="post" action="NewsApproveList.asp">
        <div class="browseby">
<%
fSql = ""
fSql = fSql & vbcrlf & "declare @user nvarchar(50)"
fSql = fSql & vbcrlf & "set @user = "& pkstr(memberId, "")
fSql = fSql & vbcrlf & "select distinct"
fSql = fSql & vbcrlf & "	 CatTreeRoot.CtRootID"
fSql = fSql & vbcrlf & "	,CatTreeRoot.CtRootName"
fSql = fSql & vbcrlf & "from NodeInfo"
fSql = fSql & vbcrlf & "INNER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
fSql = fSql & vbcrlf & "INNER JOIN KnowledgeToSubject ON CatTreeRoot.CtRootID = KnowledgeToSubject.subjectId "
fSql = fSql & vbcrlf & "WHERE (KnowledgeToSubject.status = 'Y') "
fSql = fSql & vbcrlf & "and (NodeInfo.owner = @user or exists(select UserId from infouser where UserId= @user and charindex('SysAdm',ugrpID) > 0))"


set rs = conn.execute(sql)
%>
            條件篩選依：
                主題館
            <select name="subjectId" id="subjectId" onchange="iForm.submit()" style="width:200px">
                <option value="">全部</option>
                <%
                do while not rs.eof
                    response.write "<option value='" & rs("CtRootID") & "'>" & rs("CtRootName") & "</option>"
                rs.movenext
                loop
                %>                
            </select>
            ｜審核狀態
            <select name="validate" id="validate" onchange="iForm.submit()">
                <option value="">全部</option>
                <option value="P" selected>待審核</option>
                <option value="Y">審核通過</option>                
            </select>
            ｜排序
            <select name="sort" id="sort" onchange="iForm.submit()">                
                <option value="title">文章標題</option>
                <option value="subject">主題館名稱</option>
                <option value="date">日期</option>
            </select>
            ｜新聞標題
            <input type="text" value="" name="kwTitle"  style="width:100px"/>
            <input type="submit" value="查詢" />
            <script type="text/javascript">
                iForm.subjectId.value = "<%=subjectId %>"
                iForm.validate.value = "<%=validate %>"
                iForm.sort.value = "<%=sort %>"
                iForm.kwTitle.value = "<%=kwTitle %>"
            </script>
        </div>
        <div id="Page">
            <% if cint(nowPage) <> 1 then %>
            <img src="/images/arrow_previous.gif" alt="上一頁">
            <a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>&subjectId=<%=subjectId %>&validate=<%=validate %>&sort=<%=sort %>&kwTitle=<%=server.urlencode(kwTitle)%>">上一頁</a> ，
            <% end if %>
            共<em><%=totRec%></em>筆資料，每頁顯示
            <select id="PerPage" name="pagesize" size="1" style="color: #FF0000" onchange="document.getElementById('iForm').submit()">
                <option value="15" <%if PerPageSize=15 then%> selected<%end if%>>15</option>
                <option value="30" <%if PerPageSize=30 then%> selected<%end if%>>30</option>
                <option value="50" <%if PerPageSize=50 then%> selected<%end if%>>50</option>
            </select>
            筆，目前在第
            <select name="nowPage" id="nowPageId" onchange="document.getElementById('iForm').submit()">
                <%
                for i = 1  to totPage
                    if i = cint(nowPage) then
                        response.write "<option value='"& i &"' selected>"& i &"</option>"
                    else
                        response.write "<option value='"& i &"'>"& i &"</option>"
                    end if                    
                next
                %>
            </select>
            頁
            <% if cint(nowPage)<>totPage then %>
            ，<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>&subjectId=<%=subjectId %>&validate=<%=validate %>&sort=<%=sort %>&kwTitle=<%=server.urlencode(kwTitle)%>">下一頁
                <img src="/images/arrow_next.gif" alt="下一頁"></a>
            <% end if %>
        </div>
        <table cellspacing="0" id="ListTable">
            <tr>
                <th class="eTableLable" width="30">
                    <input type="button" value="全選" class="cbutton" name="checkAllBtn" onclick="CheckAll()"></th>
                <th class="eTableLable">預覽</th>                
                <th class="eTableLable" style="width:80px">審核者</th>
                <th class="eTableLable" style="width:60px">編修日期</th>
                <th class="eTableLable" style="width:120px">單元名稱</th>
                <th class="eTableLable" style="width:120px">主題館</th>
                <th class="eTableLable">標題</th>
            </tr>
            <%
	If not RSreg.eof then   
		for i = 1 to PerPageSize
			pKey = session("myWWWSiteURL") & "/subject/ct.asp?xItem=" & RSreg("icuitem") & "&ctNode=" & RSreg("ctNodeId") & "&mp=" & RSreg("pvXdmp")
			xurl = RSreg("xURL")
%>
            <tr>
                <td class="eTableContent">
                <%if Rsreg("fCTUPublic") <> "Y" then %>
                    <input type="checkbox" name="ckbox<%=RSreg("iCUItem")%>" value="<%=RSreg("iCUItem")%>">
                <%end if%>
                </td>

                <td class="eTableContent">
                    <a href="<%=xURL%>" target="_nwGIP">view</a>
                </td>
                <td class="eTableContent">                    
                    <%
                        select case Rsreg("fCTUPublic")
                            case "Y"
                                response.write Rsreg("Signer")
                            case else
                                response.write "<font color='red'>待審核</font>"
                        end select
                    %>                    
                </td>
                <td class="eTableContent">
                    <%=RSreg("xPostDate")%>
                </td>
                <td class="eTableContent">
                    <%=RSreg("CtUnitName")%>
                </td>
                <td class="eTableContent">
                    <%=RSreg("CtRootName")%>
                </td>
                <td class="eTableContent">
                    <%=RSreg("sTitle")%>
                </td>
            </tr>
            <%
      RSreg.moveNext
      if RSreg.eof then exit for 
		next 
	end if
%>
        </table>
        <div align="center">
            <input name="button1" type="button" class="cbutton" onclick="formPass()" value="通過">
            <input name="button2" type="button" class="cbutton" onclick="formNotPass()" value="不通過">
        </div>
    </form>
</body>
</html>

<script language="javascript">
    function checkAll() {
        if (document.getElementById("checkAllBtn").value == "全選") {
            for (var i = 0; i < document.forms[0].elements.length - 1; i++) {
                if (document.forms[0].elements[i].name.substring(0, 5) == "ckbox") {
                    document.forms[0].elements[i].checked = true;
                }
            }
            document.getElementById("checkAllBtn").value = "全不選";
        }
        else {
            for (var i = 0; i < document.forms[0].elements.length - 1; i++) {
                if (document.forms[0].elements[i].name.substring(0, 5) == "ckbox") {
                    document.forms[0].elements[i].checked = false;
                }
            }
            document.getElementById("checkAllBtn").value = "全選";
        }
    }
    function formPass() {
        var pass = "";
        for (var i = 0; i < document.forms[0].elements.length - 1; i++) {
            if (document.forms[0].elements[i].name.substring(0, 5) == "ckbox") {
                if (document.forms[0].elements[i].checked) {
                    pass += document.forms[0].elements[i].value + ",";
                }
            }
        }
        if (pass == "") {
            alert("請至少選擇一項");
        }
        else {
            window.location.href = "NewsApproveList.asp?pass=Y&id=" + pass + "&nowpage=<%=nowPage%>&pagesize=<%=perpagesize%>&subjectId=<%=subjectId %>&validate=<%=validate %>&sort=<%=sort %>&kwTitle=<%=server.urlencode(kwTitle)%>";
        }
    }
    function formNotPass() {
        var notpass = "";
        for (var i = 0; i < document.forms[0].elements.length - 1; i++) {
            if (document.forms[0].elements[i].name.substring(0, 5) == "ckbox") {
                if (document.forms[0].elements[i].checked) {
                    notpass += document.forms[0].elements[i].value + ",";
                }
            }
        }
        if (notpass == "") {
            alert("請至少選擇一項");
        }
        else {
            window.location.href = "NewsApproveList.asp?notpass=Y&id=" + notpass + "&nowpage=<%=nowPage%>&pagesize=<%=perpagesize%>&subjectId=<%=subjectId %>&validate=<%=validate %>&sort=<%=sort %>&kwTitle=<%=server.urlencode(kwTitle)%>";
        }
    }
</script>

<% End Sub %>
