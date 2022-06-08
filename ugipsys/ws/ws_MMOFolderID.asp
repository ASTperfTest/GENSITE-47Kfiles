<?xml version="1.0" encoding="utf-8"?>
<divList>
<%
sub expandfrom(lngid, msgLevel, prow)
	childCount = 0
'	response.write lngid & "," & msgLevel & "," & prow & "<HR>"
    for lngrow = 0 to ubound(ARYDept, 2)
'    	response.write lngrow & ")-" & ARYDept(0, lngrow) & "," & ARYDept(1, lngrow) & "," & ARYDept(2, lngrow) & "<BR>"
      if (ARYDept(cparent, lngrow) = lngid) then
		childCount = childCount + 1
	    genlist lngrow, msgLevel+1, childCount, prow
         expandfrom ARYDept(cid,lngrow), msgLevel+1, lngrow
      end if
    next
end sub

sub genlist(introw, intnewmsglevel, cnt, prow)
	if intnewmsglevel > 0 then
	  if cnt = ARYDept(cJChild, prow) then
		leadStr = left(leadStr, intnewmsglevel-1) & "└" & mid(leadStr, intnewmsglevel+1)
	  else
		leadStr = left(leadStr, intnewmsglevel-1) & "├" & mid(leadStr, intnewmsglevel+1)
	  end if
	  if intnewmsglevel > 1 then
	  	if mid(leadStr,intnewmsglevel-1,1) = "└" then
			leadStr = left(leadStr, intnewmsglevel-2) & "　" & mid(leadStr, intnewmsglevel)
		elseif mid(leadStr,intnewmsglevel-1,1) = "├" then
			leadStr = left(leadStr, intnewmsglevel-2) & "│" & mid(leadStr, intnewmsglevel)
	  	end if
	  end if
	end if
	response.write "<row><mCode>" & ARYDept(cid,introw) & "</mCode><mValue>" & left(leadStr, intnewmsglevel) & ARYDept(cname, introw) & "</mValue></row>" & vbCRLF
	glastmsglevel = intnewmsglevel
end sub

const cid = 0
const cname = 1
const cparent = 2
const cplevel = 3
const cseq = 4
const cJChild=5
leadStr = "　　　　　　　　　　　　　"
Dim ARYDept

	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

	CtUnitID = request("CtUnitID")
	SQLM="Select mmositeId+mmofolderName as MMOFOlderPath from CtUnit C Left Join Mmofolder M ON C.mmofolderId=M.mmofolderId " & _
		"where C.ctUnitId=" & CtUnitID
	Set RSM=conn.execute(SQLM)
	if not RSM.eof then xMMOFolderID=RSM("MMOFOlderPath")
  	sqlCom="Select CASE mmofolderName WHEN 'zzz' THEN 0 ELSE mmofolderId END MMOFolderID, " & _
  		"Case MM.mmofolderParent when 'zzz' then MS.mmositeName else MM.mmofolderNameShow END MMOFolderNameShow, " & _
  		"Case mmofolderParent when 'zzz' then 0 else (Select mmofolderId from Mmofolder where mmositeId=MM.mmositeId and mmofolderName=MM.mmofolderParent) END MMOFolderParent " & _
  		",1 " & _
		",1 " & _
		",Case mmofolderName when 'zzz' then " & _
		"	(Select count(*) from Mmofolder MM2 Left Join Mmosite MS2 ON MM2.mmositeId=MS2.mmositeId where (MS2.deptId IS NULL OR MS2.deptId LIKE '" & session("deptID") & "%' OR MS2.deptId = Left('" & session("deptID") & "',Len(MS2.deptId))) and (MM2.deptId IS NULL OR MM2.deptId = Left('" & session("deptID") & "',Len(MM2.deptId)) OR MM2.deptId like '"&session("deptID")&"%') and MM2.mmofolderParent='zzz') " & _
		" else " & _
		"	(Select count(*) from Mmofolder MM2 Left Join Mmosite MS2 ON MM2.mmositeId=MS2.mmositeId where (MS2.deptId IS NULL OR MS2.deptId LIKE '" & session("deptID") & "%' OR MS2.deptId = Left('" & session("deptID") & "',Len(MS2.deptId))) and (MM2.deptId IS NULL OR MM2.deptId = Left('" & session("deptID") & "',Len(MM2.deptId)) OR MM2.deptId like '"&session("deptID")&"%') and MM2.mmofolderParent=MM.mmofolderName and MM2.mmositeId=MM.mmositeId) " & _
		" END childCount " & _
		"from Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
		"where (MS.deptId IS NULL OR MS.deptId LIKE '" & session("deptID") & "%' OR MS.deptId = Left('" & session("deptID") & "',Len(MS.deptId))) and (MM.deptId IS NULL OR  MM.deptId = LEFT('" & session("deptID") & "', Len(MM.deptId)) OR MM.deptId like '"&session("deptID")&"%') "		
	if xMMOFolderID<>"" then 
		sqlCom=sqlCom & " and CASE mmofolderName WHEN 'zzz' THEN mmofolderName ELSE MM.mmositeId + mmofolderName END like '"&xMMOFolderID&"%' "
	end if
	sqlCom=sqlCom & " order by Case mmofolderParent when 'zzz' then '' else mmofolderParent END, mmofolderId"
	set RSS = conn.execute(sqlCom)

	if not RSS.EOF then
		ARYDept = RSS.getrows(200)
		glastmsglevel = 0
		genlist 0, 0, 1, 0
	        expandfrom ARYDept(cid, 0), 0, 0
	end if
%>
</divList>
