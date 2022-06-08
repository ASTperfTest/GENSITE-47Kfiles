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
	response.write "<OPTION VALUE=" & ARYDept(cid,introw) & ">" & left(leadStr, intnewmsglevel) & ARYDept(cname, introw) & "</OPTION>" & vbCRLF
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

%>
