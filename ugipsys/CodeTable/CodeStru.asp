<%@ CodePage = 65001 %>
<% 
response.expires=0
response.charset="utf-8"
HTProgCap="代碼維護"
HTProgCode="Pn90M02"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim j
xFlag=False
if request("myid")="" then
   myid="A"
else
   myid=request("myid")
   SQL="Select codeName from CodeMetaDef where codeId=N'" & myid & "'"
   SET RSxxx=conn.execute(SQL)
   xxxcodeName=RSxxx("codeName")
end if
%>
<html>  
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches"> 
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<STYLE TYPE="text/css">
/* Outline Style Sheet */
	UL {cursor: hand; 
		color: navy;
		margin-left: 0px;
		list-style-type: none;
		list-style-image: none}
	UL UL {display: none; 
			margin-left: 0px}
	.leaf {cursor: text; color: black;
			list-style-image: none}
	.picked {background-color:pink;}
</STYLE>
</head> 
<script language="VBScript">
'	window.top.f.cols="0,*"
'	window.top.topFrame.menuimg.Src="images/x-1.gif"
</script> 
<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0" bgcolor="#FFFFFF">
 <table border="0" width="100%" cellspacing="0" cellpadding="0">     
   <tr>     
   <td width="90%" valign="bottom">     
   <DIV style="filter:Shadow(Direction=135,color='#dddddd');height:35"><font size="4">代碼維護</font></DIV></td>
</tr> 
   <tr><td>&nbsp;
  </table>
<%
rcount=0
SQL="Select Distinct C2.codeType from UgrpCode UC Left Join CodeMetaDef C2 ON UC.codeId = C2.codeId " & _
	"where  UC.ugrpId IN (" & IDStr & ") AND C2.showOrNot='Y' order by C2.codeType"
  if Instr(session("uGrpID")&",", "HTSD,") > 0 then  
	SQL="Select Distinct C2.codeType from UgrpCode UC Left Join CodeMetaDef C2 ON UC.codeId = C2.codeId " & _
		"where  1=1 order by C2.codeType"
  end if
SET RSC=conn.execute(SQL)
while not RSC.EOF 
	rcount=rcount+1
	RSC.movenext
wend
	xPos=Instr(session("uGrpID"),",")
	if xPos>0 then
		IDStr=replace(session("ugrpID"),", ","','")
		IDStr="'"&IDStr&"'"
	else
		IDStr="'" & session("ugrpID") & "'"
	end if

SQL="Select Code.codeType,Code.codeId,Code.codeName, " & _
	"(SELECT COUNT(C1.codeId) FROM (Select Distinct UC.codeId,C2.codeType from ugrpCode UC Left Join CodeMetaDef C2 ON UC.codeId = C2.codeID " & _
	"where  UC.ugrpID IN (" & IDStr & ") AND C2.ShowOrNot='Y') C1 WHERE C1.codeType = CODE.codeType) AS xCount " & _
	"from " & _
	"(Select Distinct UC.codeId,C2.codeName,C2.codeType from ugrpCode UC Left Join CodeMetaDef C2 ON UC.codeId = C2.codeID " & _
	"where  UC.ugrpID IN (" & IDStr & ") AND C2.ShowOrNot='Y') Code Left Join  CodeMetaDef C ON C.codeId=CODE.codeId " & _
	"Order by C.codeType, C.codeRank,C.codeId"
  if Instr(session("uGrpID")&",", "HTSD,") > 0 then  
	SQL="Select Code.codeType,Code.codeId,Code.codeName, " & _
		"(SELECT COUNT(C1.codeId) FROM (Select Distinct UC.codeId,C2.codeType from ugrpCode UC Left Join CodeMetaDef C2 ON UC.codeId = C2.codeID " & _
		"where  UC.ugrpID IN (" & IDStr & ") AND C2.ShowOrNot='Y') C1 WHERE C1.codeType = CODE.codeType) AS xCount " & _
		"from " & _
		"(Select Distinct UC.codeId,C2.codeName,C2.codeType from ugrpCode UC Left Join CodeMetaDef C2 ON UC.codeId = C2.codeID " & _
		"where  UC.ugrpID IN (" & IDStr & ") AND C2.ShowOrNot='Y') Code Left Join  CodeMetaDef C ON C.codeId=CODE.codeId " & _
		"Order by C.codeType, C.codeRank,C.codeId"
	SQL = "SELECT codeType, codeId, codeName, " _
		& " (SELECT count(*) FROM CodeMetaDef AS C1 WHERE C1.codeType=C.codeType) AS xCount" _
		& " FROM CodeMetaDef AS C " _
		& " ORDER BY C.codeType, C.codeRank,C.codeId"
  end if



	
'SQL="SELECT Distinct C2.codeType, C2.codeId, C2.codeName, " & _
'	"(SELECT COUNT(*) FROM CodeMetaDef AS C1 WHERE C1.codeType = C2.codeType AND C1.ShowOrNot='Y') AS xCount " & _
'	"FROM ugrpCode UC Left Join CodeMetaDef C2 ON UC.codeId=C2.codeId AND UC.ugrpID IN (" & IDStr & ") " & _
'	"Where C2.ShowOrNot='Y' ORDER BY C2.codeType DESC,convert(Int,C2.codeRank),C2.codeId "
'response.write SQL	
SET RSCode=conn.execute(SQL)
if not RSCode.EOF then%>

<DIV style="margin-left:px; line-height:20px; font-size:14px;">
<UL>
<%xx=""
  xxx=""
  p=0
  xChildCount=0    
  pivotFig="plus"
  pmFig="t"
  pmFig2=""
  leadFig = "dotsl"

  while not RSCode.eof   
'--------------資料表
              if ucase(trim(RSCode("codeType")))<>xx then
	 	 if xx <> "" and xFlag then
			response.write "</UL>" & cvCrLF
			xFlag=False
		 end if            
                 p=p+1
                 if p>1 then                 
 		    pmFig=""
		 end if
                 if rcount=p then 
		    pmFig="b"
		    leadFig="space"
  	         end if
%>
                 <LI id=s<%=RSCode("codeType")%>>
                 <img id=pms<%=RSCode("codeType")%> src=../images/tv_<%=pivotFig%>dots<%=pmFig%>.gif align=absmiddle><img src=../images/openfold.gif align=absmiddle border=0><%=RSCode("codeType")%><br>
<% 
                 xx=ucase(trim(RSCode("codeType")))
              end if
'--------------代碼
   		 if xxx = "" then
			response.write "<UL>" & cvCrLF
		 end if
                 xChildCount=xChildCount+1
 	         if RSCode("xcount")=xChildCount then
                    pmFig2="b"
		    xFlag=true
                 end if
%>
                 <LI id=s<%=RSCOde("codeId")%>>
                 <img src=../images/tv_<%=leadFig%>.gif align=absmiddle>
                 <img id=pms<%=RSCode("codeId")%> src=../images/tv_dots<%=pmFig2%>.gif align=absmiddle>
                 <a href="CodeDataDetailList.asp?codeID=<%=RSCode("codeId")%>&codeName=<%=RSCode("codeName")%>" target="article"><img src=../images/iconAttr.gif align=absmiddle border=0>&nbsp;<span id=hls<%=RSCode("codeId")%>><font size=2><%=RSCode("codeName")%></font></span></A>
<% 
                 xxx=RSCode("codeId")
             	 if RSCode("xCount")=xChildCount then
 	            pmFig2=""
  		    xChildCount=0
  		    xxx=""
		 end if
'--------------              
           RSCode.movenext
    wend
                        response.write "</UL>" & vbCrLF
'--------
%> 
</div>
<%end if
%>
</body>
</html>
<script language=javascript>
var xxxcodeName="<%=xxxcodeName%>";
function fmClubQueryonSubmit() {
  var msg;
	if (fmClubQuery.sKey.value == "") {
		msg = "請輸入欲搜尋社群的關鍵字詞";
		alert(msg);
		return false;
	}
        return true;
}


	var pickedID = "";

		function checkParent(src, dest) {
           // Search for a specific parent of the current element
           while (src!=null) {
             if (src.tagName == dest) return src;
             src = src.parentElement;
           }
           return null;
        }
		function setPlusMinusIcon(xxLIid, pORm)	{
			xIMGobj = document.all("pm"+xxLIid);
			if (xIMGobj) {
				var oImgSrc = xIMGobj.src;
				var re;
				if ("plus"==pORm) {
					re = /tv_minus/i;
					if (oImgSrc.search(re) >= 0) {
						oImgSrc = oImgSrc.replace(re, "tv_plus");
						eval("xIMGobj.src = oImgSrc");
					}
				}
				else {
					rex = /tv_plus/i;
					if (oImgSrc.search(rex) >= 0) {
						oImgSrc = oImgSrc.replace(rex, "tv_minus");
						eval("xIMGobj.src = oImgSrc");
					}
				} 
			}
		}
		function pmTaggle(xLIid) {
			xLIobj = document.all(xLIid);
			if (xLIobj && ("LI"==xLIobj.tagName))	{
	            for (var pos=0; pos<xLIobj.children.length; pos++)
					if ("UL"==xLIobj.children[pos].tagName) break;
	            if (pos==xLIobj.children.length) return;
				el = xLIobj.children[pos];
				if (""==el.style.display) {
					el.style.display = "block";
					setPlusMinusIcon(xLIid,"minus");
				} else {
					el.style.display = "";
					setPlusMinusIcon(xLIid,"plus");
				}
			}
		}
        function outline() {     
           // Expand or collapse if a list item is clicked.
           var open = event.srcElement;
			if (("A"==open.parentElement.tagName) || "SPAN"==open.tagName) {
				var el = checkParent(open, "LI");
				if (null!=el)	showPicked("hl"+el.id);
				return;
			}
			if ("LI" != open.tagName && "IMG" !=open.tagName)	return;
			if ("LI" == open.tagName)	pmTaggle(open.id);
			if ("IMG" == open.tagName && open.parentElement.tagName!="LI")	return;
			if ("IMG" == open.tagName && open.parentElement.tagName=="LI") {
				pmTaggle(open.parentElement.id);
			}
          event.cancelBubble = true;
		}

		function expand(pl, el) {
			if ("UL"==el.tagName) { 
				el.style.display = "block"; 
			}
			if ("IMG"==el.tagName) {
				var oImgSrc = el.src
				var re, rex;
				rex = /tv_plus/i;
				if (oImgSrc.search(rex) >= 0) {
					oImgSrc = oImgSrc.replace(rex, "tv_minus");
					eval("el.src = oImgSrc");
				}
			}
			for (var pos=0; pos<el.children.length; pos++) {
				expand(el, el.children[pos]);
			}
		}
        function outall() {     
           // Expand or collapse if a list item is clicked.
           var open = event.srcElement;
			if ("IMG" != open.tagName || open.parentElement.tagName!="LI")	return;

				var oImgSrc = open.src
				var re, rex;
					rex = /tv_plus/i;
					if (oImgSrc.search(rex) >= 0) {
						oImgSrc = oImgSrc.replace(rex, "tv_minus");
						eval("open.src = oImgSrc");
					}

           // Make sure clicked inside an LI. This test allows rich HTML inside lists.
           var el = checkParent(open, "LI");
           if (null!=el) {
             var pos = 0;
             // Search for a nested list
             for (var pos=0; pos<el.children.length; pos++) {
               if ("UL"==el.children[pos].tagName) break;
            }
            if (pos==el.children.length) return;
          } else return;
		var pl = el
          el = el.children[pos];
          if ("UL"==el.tagName) {
            // Expand or Collapse nested list
			expand(pl, el);
        }
          event.cancelBubble = true;
        }

		function showPicked(xLIid) {
			if (""!=pickedID) 	document.all(pickedID).className="";
			pickedID = xLIid;
			document.all(pickedID).className="picked";
		}
		function xfocusIt(xLIid) {
			if ("sA"==xLIid)  return;
			px = document.all(xLIid);
			x = px.parentElement;
			while (x!=null) {
				if (("LI"==x.tagName) && ("sA"!=x.id))	setPlusMinusIcon(x.id,"minus");
				if ("UL"==x.tagName)	x.style.display = "block";
				x = x.parentElement;
			}
			px.scrollIntoView();
			xa = document.all("hl"+xLIid);
			showPicked(xa.id);
		}

        document.onclick = outline;

		document.ondblclick = outall;

//		sA.click();
		xfocusIt("s<%=myid%>");
		if (xxxcodeName != "") {
			parent.frames(1).navigate("CodeDataDetailList.asp?codeID=<%=myid%>&codeName=<%=xxxcodeName%>");
		}
</script>
