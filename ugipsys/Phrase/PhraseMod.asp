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
'設定詞彙程式路徑
PhraseModUrl = ""
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
	
		Set Conn = Server.CreateObject("ADODB.Connection")	
		Conn.Open session("ODBCDSN")
	
	if  request("Create")="true" then
	
	%>
	
	
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="/css/form.css" rel="stylesheet" type="text/css">
    <link href="/css/layout.css" rel="stylesheet" type="text/css">
    <title>資料管理／資料上稿</title>
</head>
<body

    <div id="FuncName">
        <h1>主題館詞彙解釋新增</h1>
        <div id="ClearFloat">
        </div>
    </div>
    <form id="Form1" method="POST" name="reg" action="../Phrase/PhraseModAct.asp?CtRootID=<%=request("CtRootID")%>&mode=Create" onsubmit="formModSubmit()">
	
    <table cellspacing="0">
        <tr>
            <td class="Label" align="right"><span class="Must">*</span>詞彙標題</td>
            <td class="eTableContent">
                <input name="htx_stitle" size="50" value="" maxlength="30">
            </td>
        </tr>
        <tr>
            <td class="Label" align="right">內文</td>
            <td class="eTableContent">
                <textarea name="htx_xbody" rows="8" cols="60"></textarea><br />
            </td>
        </tr>
        <tr>
            <td class="Label" align="right"><span class="Must">*</span>是否公開</td>
            <td class="eTableContent">
                <select id="htx_fctupublic" name="htx_fctupublic" size="1">
                    <option value="">請選擇</option>
                    <option value="Y">公開</option>
                    <option value="N">不公開</option>
                </select>
            </td>
        </tr>
    </table>

    <input type="submit" value="新增存檔" class="cbutton" >  
    <input type="reset" value="重　填" class="cbutton">
    <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
    </form>
	
	<%
		
		else
			xsql = "select * from SubjectPhrase where rowID='" & request("rowID")& "'"
			set xsqllist = conn.execute (xsql)	
		
%>


<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="/css/form.css" rel="stylesheet" type="text/css">
    <link href="/css/layout.css" rel="stylesheet" type="text/css">
    <title>資料管理／資料上稿</title>
</head>
<body>

    <div id="FuncName">
        <h1>主題館詞彙解釋維護</h1>
        <div id="ClearFloat">
        </div>
    </div>
    <form id="Form1" method="POST" name="reg" action="../Phrase/PhraseModAct.asp?CtRootID=<%=request("CtRootID")%>&rowID=<%=xsqllist("rowID")%>&mode=Modify" onsubmit="formModSubmit()">
	
	<input type="hidden" name="modify_rowID" value="<%=xsqllist("rowID")%>">	
	<input type="hidden" name="modify_Phrase" value="<%=trim(xsqllist("Phrase"))%>">
	
    <table cellspacing="0">
        <tr>
            <td class="Label" align="right"><span class="Must">*</span>詞彙標題</td>
            <td class="eTableContent">
                <input name="htx_stitle" size="50" value="<%=trim(xsqllist("Phrase"))%>">
            </td>
        </tr>
        <tr>
            <td class="Label" align="right">內文</td>
            <td class="eTableContent">
                <textarea name="htx_xbody" rows="8" cols="60"><%=xsqllist("Content")%></textarea><br />
            </td>
        </tr>
        <tr>
            <td class="Label" align="right"><span class="Must">*</span>是否公開</td>
            <td class="eTableContent">
                <select id="htx_fctupublic" name="htx_fctupublic" size="1">
                    <option value="">請選擇</option>
                    <option value="Y" <% if xsqllist("FCTUPublic") = "Y" then response.write "selected" end if %>>公開</option>
                    <option value="N" <% if xsqllist("FCTUPublic") = "N" then response.write"selected" end if %>>不公開</option>
                </select>
            </td>
        </tr>
    </table>
    <script type="text/javascript">
        function delRow(rowID) {
            if (confirm("確認要刪除此筆資料嗎?")) {
                location.href = 'PhraseModAct.asp?CtRootID=<%=request("CtRootID")%>&rowID=' + rowID + '&mode=del';
            }
        }        
    </script>    
	<input type="submit" value="編輯存檔" class="cbutton">    
    <input type=button value="刪　除" class="cbutton" onClick="delRow(<%=xsqllist("rowID")%>)">
	<input type=button value ="重　填" class="cbutton" >
    <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
</body>
</html>

<%end if  %>
 

 <script language="vbscript">
      sub formModSubmit()
    
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
  
	IF reg.htx_stitle.value = Empty Then 
		MsgBox replace(nMsg,"{0}","詞彙標題"), 64, "Sorry!"
		reg.htx_stitle.focus
		window.event.returnvalue=false 
		exit sub
	END IF

	IF  Len(reg.htx_xbody.value)>500 Then 
		MsgBox "請勿超過五百字"
		reg.htx_stitle.focus
		window.event.returnvalue=false 
		exit sub
	END IF
	
	IF reg.htx_fctupublic.value = Empty Then 
		MsgBox replace(nMsg,"{0}","公開選項"), 64, "Sorry!"
		reg.htx_fctupublic.focus
	window.event.returnvalue=false 		
	    exit sub
	END IF

   end sub
</script>
 
 <% End Sub %>