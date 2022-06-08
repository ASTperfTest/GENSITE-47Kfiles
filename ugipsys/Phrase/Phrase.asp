<%@  codepage="65001" %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
'設定詞彙程式路徑
PhraseModUrl = "Phrase.asp"
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
<!--#include virtual = "/inc/ReplaceAndFindKeyword.inc" -->
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
			document.location.href="";
		<% else %>
			history.go(-1)
		<% end if %>
		
	</script>

</body>
</html>
<% End sub 
Sub showForm 
Phrase="SELECT * FROM SubjectPhrase where CtRootID=" & request("mp")
set Phrase_List = conn.execute(Phrase)
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>詞彙維護</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="/Project0516/css/list.css" rel="stylesheet" type="text/css" />
    <link href="/css/layout.css" rel="stylesheet" type="text/css" />
    <link href="/css/theme.css" rel="stylesheet" type="text/css" />
    <script src="/js/jquery.js" type="text/javascript"></script>
    <script src="/js/jtip.js" type="text/javascript"></script>

    <style>
        .h1
        {
            text-align: left;
            padding-left: 5px;
        }
        .h3
        {
            width: 60px;
            text-align: left;
            padding-left: 5px;
        }
    </style>
	
	    <style type="text/css" media="all">
        @import "/css/global.css";
		
        .style1
        {
            height: 23px;
        }
    </style>

</head>
<body>

  <form id="Formlist" method="POST" name="reg" action="../Phrase/PhraseMod.asp?CtRootID=<%=request("mp")%>&Create=true">
    <div id="FuncName">
        <h1>主題館詞彙解釋</h1>
        <div id="ClearFloat">
        </div>
    </div>

    <div>
        <table class="ListTable" cellspacing="0" rules="all" border="1" id="GridView1" style="border-collapse: collapse;">
            <tr>
                <th scope="col" class="h1">詞彙標題</th>
                <th scope="col" class="h3">瀏覽</th>
                <th scope="col" class="h3">是否公開</th>
                <th scope="col" class="h3">張貼日</th>
            </tr>
			<% while not Phrase_List.eof %>
            <tr>
                <td>
                    <a href="../Phrase/PhraseMod.asp?CtRootID=<%=Phrase_List("CtRootID")%>&rowID=<%=Phrase_list("rowId")%>"><%=server.htmlencode(Phrase_list("Phrase"))%></a>
                </td>
                <td><%=ReplaceAndFindKeyword(Phrase_list("Phrase"))%></td>
				
                <td><%=Phrase_list("fCTUPublic")%></td>
				
                <td><%=Phrase_list("CreationDT")%></td>
				
            </tr>
				<% 
				   Phrase_list.moveNext 
				   wend
			
				%>
            
        </table>
    </div>
	<input type="submit" value="新增">


    </form>
</body>
</html>
<% end sub %>


