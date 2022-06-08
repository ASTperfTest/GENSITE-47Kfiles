<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GC1AP3" %>
<!--#include virtual = "/inc/server.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function
function blen(xs)              '檢測中英文夾雜字串實際長度
  xl = len(xs)
  for i=1 to len(xs)
	if asc(mid(xs,i,1))<0 then xl = xl + 1
  next
  blen = xl
end function

sub traverseFolder(MMOFolderName,TreeID,MMOSiteID)
	xMMOSiteID=MMOSiteID
	xTreeID=TreeID
	xMMOFolderName=MMOFolderName
    	SQL="Select M.MMOFolderName,M.MMOFolderDesc,M.MMOFolderNameShow " & _
    		"from MMOFolder M " & _ 
    		"where M.MMOSiteID='"&xMMOSiteID&"' and SUBSTRING(M.MMOFolderName, 1, "&len(xMMOFolderName)&")='"&xMMOFolderName&"' and M.MMOFolderParent='"& xMMOFolderName & "' " & _
    		" and (M.deptID is null or M.deptID LIKE '" & session("deptID") & "%' OR M.deptID = Left('" & session("deptID") & "',Len(M.deptID))) order by M.MMOFolderID"
	Set RSD=conn2.execute(SQL)  
	if not RSD.EOF then
	    while not RSD.EOF
		yMMOFolderName=RSD(0) 
		yMMOFolderDesc=RSD(1) 
		if left(yMMOFolderName,1)="/" then yyMMOFolderName = mid(yMMOFolderName,2)	'---為跑後面原有程式,將目錄路徑左方/去除
    		SQLCount="Select count(M2.MMOFolderID) from MMOFolder M2 where (M2.deptID is null or M2.deptID LIKE '" & session("deptID") & "%' OR M2.deptID = Left('" & session("deptID") & "',Len(M2.deptID))) and M2.MMOSiteID='"&xMMOSiteID&"' and SUBSTRING(M2.MMOFolderName, 1, "&len(yMMOFolderName)&")='"&yMMOFolderName&"' and M2.MMOFolderParent='"& yMMOFolderName & "'"
    	'response.write SQLCount
		SET RSCoount=conn2.execute(SQLCount)	    
		yChildCount=RSCoount(0)
		if yChildCount>=1 then
			Response.Write replace(xMMOSiteID&yMMOFolderName,"/","") & " = insFld("&xTreeID&", gFld(""ForumToc"", """& RSD("MMOFolderNameShow") &""", ""MMOEditFolder.asp?MMOSiteID="&xMMOSiteID&"&xpath="&yyMMOFolderName&""","""&yMMOFolderDesc&"""))" & vbcrlf		
			traverseFolder yMMOFolderName, replace(xMMOSiteID&yMMOFolderName,"/",""),xMMOSiteID
		else
			Response.Write "insDoc("&xTreeID&", gLnk(""ForumToc"", """& RSD("MMOFolderNameShow")&""", ""MMOEditFolder.asp?MMOSiteID="&xMMOSiteID&"&xpath="&yyMMOFolderName&""","""&yMMOFolderDesc&"""))" & vbcrlf			
		end if
		RSD.movenext
	    wend	
	end if	
end sub
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language=javascript>
 var OpenFolder = "/imageFolder/ftv2folderclosed.gif"
 var CloseFolder = "/imageFolder/ftv2folderopen.gif"
 var DocImage = "/imageFolder/ftv2folderclosed.gif"
</script>
<script src="/imageFolder/FolderTree.js"></script>
<base target="Catalogue">
<title></title>
<style>
.FolderStyle { text-decoration: none; color:#000000}
</style>
</head><body leftmargin="3" topmargin="3">
<script language=javascript>
<%
Response.write "treeRoot = gFld(""ForumToc"", ""多媒體物件　　　　　<A href='MMOSiteAdd.asp' target='ForumToc'>新增站台</A>"", """",""多媒體物件根目錄"")" & vbcrlf
SQLMMO="Select M.* " & _
	",(Select count(*) from MMOFolder where MMOSiteID=M.MMOSiteID and (deptID is null or deptID LIKE '" & session("deptID") & "%' OR deptID = Left('" & session("deptID") & "',Len(deptID)))) MMOSiteChildCount " & _
	"from MMOSite M  " & _
	"Where (deptID is null or deptID LIKE '" & session("deptID") & "%' OR deptID = Left('" & session("deptID") & "',Len(deptID))) Order by M.MMOSiteID"
Set RSMMO=conn.execute(SQLMMO)
'response.write SQLMMO & vbcrlf

if not RSMMO.EOF then

    while not RSMMO.EOF

 if RSMMO("MMOSiteChildCount") > 0 then

    	Set conn2 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'    	conn2.Open session("ODBCDSN")
'Set conn2 = Server.CreateObject("HyWebDB3.dbExecute")
conn2.ConnectionString = session("ODBCDSN")
conn2.ConnectionTimeout=0
conn2.CursorLocation = 3
conn2.open
'----------HyWeb GIP DB CONNECTION PATCH----------

    	SQL="Select MS.MMOSiteID,MS.MMOSiteName " & _
    		",(Select count(M2.MMOFolderID) from MMOFolder M2 where M2.MMOSiteID=MS.MMOSiteID and (M2.deptID is null or M2.deptID LIKE '" & session("deptID") & "%' OR M2.deptID = Left('" & session("deptID") & "',Len(M2.deptID)))) ChildCount " & _
    		",M.MMOFolderName " & _
    		"from MMOSite MS Left Join MMOFolder M ON MS.MMOSiteID=M.MMOSiteID " & _ 
    		"where (M.deptID is null or M.deptID LIKE '" & session("deptID") & "%' OR M.deptID = Left('" & session("deptID") & "',Len(M.deptID))) and MS.MMOSiteID='"&RSMMO("MMOSiteID")&"' and SUBSTRING(M.MMOFolderName, 2, LEN(M.MMOFolderName) - 1)=''"
	Set RSF=conn2.execute(SQL) 

	if not RSF.EOF then
	  while not RSF.EOF   
	    zMMOSiteID=RSF("MMOSiteID")
	    zMMOSiteName=RSF("MMOSiteName")
	    zChildCount=RSF("ChildCount")	
	    zMMOFolderName=RSF("MMOFolderName")  
	    if zChildCount > 1 then
		Response.Write zMMOSiteID & " = insFld(treeRoot, gFld(""ForumToc"", """& zMMOSiteName &""", ""MMOSiteEdit.asp?MMOSiteID="&zMMOSiteID&""","""&zMMOSiteID&"""))" & vbcrlf		
		traverseFolder zMMOFolderName, zMMOSiteID ,RSMMO("MMOSiteID")
	    else
		Response.Write "insDoc(treeRoot, gLnk(""ForumToc"", """& zMMOSiteName &""", ""MMOSiteEdit.asp?MMOSiteID="&zMMOSiteID&""","""&zMMOSiteID&"""))" & vbcrlf	
	    end if
	    RSF.movenext
	  wend
	end if	
      else

	Response.Write "insDoc(treeRoot, gLnk(""ForumToc"", """& RSMMO("MMOSiteName") &""", ""MMOSiteEdit.asp?MMOSiteID="&RSMMO("MMOSiteID")&""","""&RSMMO("MMOSiteID")&"""))" & vbcrlf	
      end if
      RSMMO.movenext
    wend

end if
NowCount = 0


%>
    initializeDocument();
</script>
</body></html>

