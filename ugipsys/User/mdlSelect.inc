<%response.expires=0%>
<html>
<%

const cmaxmessagelevel = 10
myid = "A"
xxKind = "D"

const cid = 1
const cname = 2
const ename = 2
const cparent = 3
const ckind = 4
const cplevel = 5
const cseq = 6
const cChild = 7
const cJChild = 15
const cMaxForum = 2000
'ReDim ARYForum(cForumAttributes,cMaxForum)

%>
 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"> 
<meta http-equiv="Pragma" content="no-cache"> 
<title>Products</title> 
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
	.picked {background-color:pink; cursor:hand;}
</STYLE>
 
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head> 
 
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【<%=HTProgCap%>授權】&nbsp;&nbsp;&nbsp;&nbsp;授權UserID:[<%=UID%>]</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td colspan=2 width=80% height=230 valign=top>        
<form name=pform method=post><span id=OKB style="visibility:hidden;"><input type=submit value=" OK " style="position:absolute;top:50pt;left:120pt"></span>
<DIV style="margin-left:-0px; line-height:20px; font-size:14px;">
<%
SQLCom = "SELECT D.*,(Select Count(*) from deptOrg AS D2 where D.VerID=D2.VerID AND D.deptID=D2.parent AND kind='J') FROM deptOrg AS D WHERE D.verID=(Select VerID from OrgVersion where VerType='ORG' AND ExpireDate IS NULL) ORDER BY D.plevel, D.kind desc, D.seq"
set RS= conn.execute(SQLCom)

	leadString = "　　　　　　　　　　　　　　　"

if not rs.eof then
   ARYForum = rs.getrows(cMaxForum)

'	response.write "<TABLE border>"
'	for xi = 0 to ubound(ARYForum,2)
'		response.write "<TR>"
'		for xj = 0 to ubound(ARYForum,1)
'			response.write "<TD>" & ARYForum(xj,xi)
'		next
'	next
'	response.write "</TABLE>"

   glastmsglevel = 0
		xgenlist 0, 1, true
		xexpandfrom 0, 1
	for intIndex = 1 to glastmsglevel
		response.write "</ul>"
	next
else
   response.write "<b>目前還沒有任何資料!!</b>"
end if

''conn.close

sub xexpandfrom (pidx, xlevel)
	xcid = ARYForum(cid, pidx)
	childCount = 0
	for i = 0 to ubound(ARYForum,2)
		if ARYForum(cparent,i) = xcid then
			childCount = childCount + 1
		   if ARYForum(ckind,i)="d" then
			if childCount = ARYForum(cChild, pidx) then
				call xgenlist (i, xlevel+1, true)
			else
				call xgenlist (i, xlevel+1, false)
			end if
			xexpandfrom i, xlevel + 1
		    end if
		end if
	next
end sub

sub xgenlist(xidx, xlevel, lastChild)
  Dim intIndex

'	response.write xidx & "," & xlevel & "," & lastChild & "<HR>"
'	exit sub
	
	intnewmsglevel = xlevel
  for intIndex = glastmsglevel to intnewmsglevel - 1
    response.write "<ul type=square>" & vbcrlf
  next
  for intIndex = intNewMsgLevel to glastmsglevel - 1
    response.write "</ul>"
  next
  if intnewmsglevel > 0 then

	introw = xidx
	if (ARYForum(cChild, introw) - ARYForum(cJChild, introw)) > 0 then
		strLI = "<li id=s" & ARYForum(cid, introw) & " style=""background-color:#" & hex(15-ARYforum(cpLevel,introw)*3) & hex(15-ARYforum(cpLevel,introw)*2) & "C"">"
		strLI = "<li id=s" & ARYForum(cid, introw)
		xxKind=" myKind=F"
		iconPrefix = "tv_plusdots"
		iconFolder = "iconXML.gif"
		xxanchorURL = "elemEdit.asp"
		pmTag = " id=pms" &  ARYForum(cid, introw)
	else
		iconPrefix = "tv_dots"
		strLI = "<li class=leaf id=s" & ARYForum(cid, introw) & " style=""background-color:white"">"
		strLI = "<li class=leaf id=s" & ARYForum(cid, introw)
		if ARYForum(cKind, introw) = "A" then
			iconFolder = "iconAttr.gif"
			xxanchorURL = "attrEdit.asp"
			xxKind=" myKind=A"
		else
			xxKind=" myKind=F"
			iconFolder = "iconXML.gif"
			xxanchorURL = "elemEdit.asp"
		end if
		pmTag = ""
	end if
	if xlevel > 0 then 
		xxStr = left(leadString,xlevel-1)
		xxStr = replace(xxStr,"│", "<img src=../images/tv_dotsl.gif align=absmiddle>")
		xxStr = replace(xxStr,"　", "<img src=../images/tv_space.gif align=absmiddle>")
		response.write strLI & ">" & xxStr

		if lastChild then
			response.write "<img" & pmTag & " src=../images/" & iconPrefix & "b.gif align=absmiddle>"
			leadString = left(leadString,xlevel-1) & "　" & mid(leadString,xlevel)
		else
			response.write "<img" & pmTag & " src=../images/" & iconPrefix & ".gif align=absmiddle>"
			leadString = left(leadString,xlevel-1) & "│" & mid(leadString,xlevel)
		end if
	else
		response.write "<LI id=s1><img id=pms1 src=../images/tv_space.gif width=0>"

	end if
'	response.write "<a href=" & xxanchorURL & "?nodeID=" & ARYForum(cid, introw) & "><img src=../images/" & iconFolder & " align=absmiddle border=0>&nbsp;<span id=hls" & ARYForum(cid, introw) &">" & ARYForum(1,xidx) & "</span></A>" & vbCRLF
	response.write "<img src=../images/" & iconFolder & " align=absmiddle border=0>&nbsp;"
	if ARYForum(ckind, introw) = "d" then
		response.write "<INPUT TYPE=checkbox NAME=mdlSelected" & ARYForum(cid, introw) & ">"
	end if
	response.write "&nbsp;<span id=hls" & ARYForum(cid, introw) & xxKind & " title=""" & ARYForum(cname,xidx) & """>" & ARYForum(ename,xidx) & "</span>" & vbCRLF
  end if
  glastmsglevel = intnewmsglevel
end sub

%>
</DIV></form>
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Copyright     
      &copy;2000 Chung Shan Medical and Dental College.<br>   
      Best viewed with Internet Explorer 4.0+ , Screen Size set to 800x600.</td>                                         
  </tr> 
</table> 
</body> 

<SCRIPT LANGUAGE="JavaScript">

	var pickedID = "";
	var overID = "";

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
//			if (("A"==open.parentElement.tagName) || "SPAN"==open.tagName) {
			if ("SPAN"==open.tagName) {
//				alert (open.myKind);
				if ("F"==open.myKind) {
//					alert (open.id);
//					alert (open.innerText);
				    window.returnValue=open.id + ";;;" + open.innerText;
				    window.close();
				}
//				var el = checkParent(open, "LI");
//				if (null!=el)	showPicked("hl"+el.id);
//				return;
			}
			if ("LI" != open.tagName && "IMG" !=open.tagName)	return;
			if ("LI" == open.tagName)	pmTaggle(open.id);
			if ("IMG" == open.tagName && open.parentElement.tagName!="LI")	return;
			if ("IMG" == open.tagName && open.parentElement.tagName=="LI") {
				pmTaggle(open.parentElement.id);
			}
          event.cancelBubble = true;
		}

        function xOver() {     
           var open = event.srcElement;
			if ("SPAN"==open.tagName) {
				if ("F"==open.myKind) {
					if (""!=overID) 	document.all(overID).className="";
					overID = open.id;
					document.all(overID).className="picked";
				}
			}
//          event.cancelBubble = true;
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
			overID = xLIid;
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
        document.onmouseover = xOver;

		document.ondblclick = outall;

		sA.click();
		xfocusIt("s<%=myid%>");
</SCRIPT>
</html> 
