<!--#Include virtual = "/inc/client.inc" -->
<%
function AlertAndGoSubjectList(msg, mp)    
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
    <script type="text/javascript">
        alert('<%=msg %>');
        <%if isnumeric(mp) and mp<>"" then %>
            window.location.href = "/subject/mp.asp?mp=<%=mp %>"
        <%else%>
            window.location.href = "/subjectlist.aspx"
        <%end if%>
    </script>
</body>    
</html>    
<%
response.Status = "403 Forbidden"
response.End
end function

function CheckPoint_BeforeLoadXML(mp, ctNode, xItem)

    if session("backEndSetup") <> "" then
        '從後台瀏覽，不檢查目錄、文章是否開放
        exit function        
    end if

    
    if ctNode = "" then    
        if not isnumeric(mp) or mp="" then
            AlertAndGoSubjectList "主題館不存在或未開放!!", mp
        end if    
            
        sqlstr ="select * from CatTreeRoot where ctRootId = " & mp    
        sqlstr = sqlstr & vbcrlf & " and inUse ='Y'"
            
        set rs = conn.execute(sqlstr)
        if rs.eof then 
            AlertAndGoSubjectList "主題館不存在或未開放!!", ""
            conn.close()
            set conn = nothing
        end if
    else
        if not isnumeric(ctNode) then
            AlertAndGoSubjectList "目錄不存在或未開放!!", mp
        end if  
                
        sqlstr=""
        sqlstr = sqlstr & vbcrlf & " declare @ctNodeId int "
        sqlstr = sqlstr & vbcrlf & " set @ctNodeId = " & ctNode
        sqlstr = sqlstr & vbcrlf & " with node_Tree(ctRootId, dataParent, ctNodeId, inUse) as"
        sqlstr = sqlstr & vbcrlf & " ("
        sqlstr = sqlstr & vbcrlf & " 	select "
        sqlstr = sqlstr & vbcrlf & " 		 ctRootId"
        sqlstr = sqlstr & vbcrlf & " 		,dataParent"
        sqlstr = sqlstr & vbcrlf & " 		,ctNodeId"
        sqlstr = sqlstr & vbcrlf & " 		,inUse"
        sqlstr = sqlstr & vbcrlf & " 	from CatTreeNode where ctNodeId = @ctNodeId	"
        sqlstr = sqlstr & vbcrlf & " 	union all"
        sqlstr = sqlstr & vbcrlf & " 	select "
        sqlstr = sqlstr & vbcrlf & " 		 CatTreeNode.ctRootId"
        sqlstr = sqlstr & vbcrlf & " 		,CatTreeNode.dataParent"
        sqlstr = sqlstr & vbcrlf & " 		,CatTreeNode.ctNodeId"
        sqlstr = sqlstr & vbcrlf & " 		,CatTreeNode.inUse"
        sqlstr = sqlstr & vbcrlf & " 	from node_Tree"
        sqlstr = sqlstr & vbcrlf & " 	inner join CatTreeNode on node_Tree.dataParent = CatTreeNode.ctNodeId"
        sqlstr = sqlstr & vbcrlf & " )"
        sqlstr = sqlstr & vbcrlf & " select * from node_Tree where inUse != 'Y'"
        sqlstr = sqlstr & vbcrlf & " union"
        sqlstr = sqlstr & vbcrlf & " select" 
        sqlstr = sqlstr & vbcrlf & " 	 ctRootId"
        sqlstr = sqlstr & vbcrlf & " 	,0 as dataParent"
        sqlstr = sqlstr & vbcrlf & " 	,0 as ctNodeId"
        sqlstr = sqlstr & vbcrlf & " 	,inUse	"
        sqlstr = sqlstr & vbcrlf & " from CatTreeRoot "
        sqlstr = sqlstr & vbcrlf & " where inUse != 'Y' and ctRootId in (select ctRootId from CatTreeNode where ctNodeId = @ctNodeId);"

        set rs = conn.execute(sqlstr)
        if not rs.eof then 
            AlertAndGoSubjectList "目錄不存在或未開放!!", mp
            conn.close()
            set conn = nothing
        end if
    end if
    
    if isnumeric(xItem) and xItem<>"" then
        sqlstr ="select icuitem from CuDTGeneric where fCTUPublic = 'Y' and icuitem = " & xItem
        set rs = conn.execute(sqlstr)
        if rs.eof then 
            AlertAndGoSubjectList "文件不存在或未開放!!", mp
            conn.close()
            set conn = nothing
        end if        
        
    end if
    
end function
%>