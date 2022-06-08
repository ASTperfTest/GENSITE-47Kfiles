﻿<?xml version="1.0"  encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:hyweb="urn:gip-hyweb-com"
xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone= "yes"/>


<xsl:template match="tocTree">

<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches"/> 
<base target="ForumToc"/>
<title>tocTree</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css"/>
<STYLE TYPE="text/css">
/* Outline Style Sheet */
body {
	background: #F1F5E9 url(../images/tree_bg.gif) repeat-x;
	margin: 10px;
	padding: 0px;
}
ul {
	cursor: hand; 
	color: navy;
	margin-left: 0px;
	list-style-type: none;
	list-style-image: none;
	padding: 0;
}
ul ul {
	display: none; 
	margin-left: 0px;
}
.leaf {
	cursor: text;
	color: black;
	list-style-image: none
}
.picked {
	background-color: pink;
}</STYLE>
</head>

<body>
<script language="javascript">
<xsl:text disable-output-escaping="yes">
<![CDATA[
	var pickedID = "";

		function checkParent(src, dest) {
           // Search for a specific parent of the current element
           while (src!=null) {
             if (src.tagName == dest) return src;
             src = src.parentNode;
           }
           return null;
        }
		function setPlusMinusIcon(xxLIid, pORm)	{
			xIMGobj = document.getElementById("pm"+xxLIid);
			if (xIMGobj) {
				var oImgSrc = xIMGobj.src;
				var re;
				if ("plus"==pORm) {
					re = /fclose/i;
					if (oImgSrc.search(re) >= 0) {
						oImgSrc = oImgSrc.replace(re, "fopen");
						xIMGobj.src = oImgSrc;
					}
				}
				else {
					rex = /fopen/i;
					if (oImgSrc.search(rex) >= 0) {
						oImgSrc = oImgSrc.replace(rex, "fclose");
						xIMGobj.src = oImgSrc;
					}
				} 
			}
		}
		function pmTaggle(xLIid) {
			xLIobj = document.getElementById(xLIid);
			if (xLIobj && ("LI"==xLIobj.tagName))	{
	            for (var pos=0; pos<xLIobj.childNodes.length; pos++)
					if ("UL"==xLIobj.childNodes[pos].tagName) break;
	            if (pos==xLIobj.childNodes.length) return;
				el = xLIobj.childNodes[pos];
				if (""==el.style.display) {
					el.style.display = "block";
					setPlusMinusIcon(xLIid,"minus");
				} else {
					el.style.display = "";
					setPlusMinusIcon(xLIid,"plus");
				}
			}
		}
        function outline(event) {     
           event = event || window.event;
           // Expand or collapse if a list item is clicked.
           var open = event.target || event.srcElement;
//           alert (open.outerHTML);
			if (("A"==open.parentNode.tagName) || "SPAN"==open.tagName) {
				var el = checkParent(open, "LI");
				if (null!=el)	showPicked("hl"+el.id);
				return;
			}
			
			/*
			if ("LI" != open.tagName && "IMG" !=open.tagName)	return;
			if ("LI" == open.tagName)	pmTaggle(open.id);
			if ("IMG" == open.tagName && open.parentNode.tagName!="LI")	return;
			if ("IMG" == open.tagName && open.parentNode.tagName=="LI") {
				pmTaggle(open.parentNode.id);
			}
			*/
			
			if("LI" == open.tagName) {
				pmTaggle(open.id);
			}
			else if (("IMG" == open.tagName || "A" == open.tagName) && open.parentNode.tagName == "LI") {
				pmTaggle(open.parentNode.id);
			}
			
          event.cancelBubble = true;
          if (event.stopPropagation)
            event.stopPropagation();
		}

		function expand(pl, el) {
			if ("UL"==el.tagName) { 
				el.style.display = "block"; 
			}
			if ("IMG"==el.tagName) {
				var oImgSrc = el.src;
				var re, rex;
				rex = /fopen/i;
				if (oImgSrc.search(rex) >= 0) {
					oImgSrc = oImgSrc.replace(rex, "fclose");
					el.src = oImgSrc;
				}
			}
			for (var pos=0; pos<el.childNodes.length; pos++) {
				expand(el, el.childNodes[pos]);
			}
		}
        function outall(event) {     
           event = event || window.event;
           // Expand or collapse if a list item is clicked.
           var open = event.target || event.srcElement;
			if ("IMG" != open.tagName || open.parentNode.tagName!="LI")	return;

				var oImgSrc = open.src;
				var re, rex;
					rex = /fopen/i;
					if (oImgSrc.search(rex) >= 0) {
						oImgSrc = oImgSrc.replace(rex, "fclose");
						open.src = oImgSrc;
					}

           // Make sure clicked inside an LI. This test allows rich HTML inside lists.
           var el = checkParent(open, "LI");
           if (null!=el) {
             var pos = 0;
             // Search for a nested list
             for (var pos=0; pos<el.parentNode.length; pos++) {
               if ("UL"==el.parentNode[pos].tagName) break;
            }
            if (pos==el.parentNode.length) return;
          } else return;
		var pl = el
          el = el.parentNode[pos];
          if ("UL"==el.tagName) {
            // Expand or Collapse nested list
			expand(pl, el);
        }
          event.cancelBubble = true;
          if (event.stopPropagation)
            event.stopPropagation();
        }

		function showPicked(xLIid) {
			var picked = document.getElementById(pickedID);
			if (""!=pickedID) 	picked.className="";
			pickedID = xLIid;
			picked.className="picked";
		}
		function xfocusIt(xLIid) {
			if ("sA"==xLIid)  return;
			px = document.getElementById(xLIid);
			x = px.parentNode;
			while (x!=null) {
				if (("LI"==x.tagName) && ("sA"!=x.id))	setPlusMinusIcon(x.id,"minus");
				if ("UL"==x.tagName)	x.style.display = "block";
				x = x.parentNode;
			}
			px.scrollIntoView();
			xa = document.getElementById("hl"+xLIid);
			showPicked(xa.id);
		}

        document.onclick = outline;

		document.ondblclick = outall;

//		sA.click();
//		xfocusIt("sA");
//		if (xxxCodeName != "") {
//			parent.frames(1).navigate("CodeDataDetailList.asp?codeID=A&CodeName=");
//		}
]]></xsl:text>
</script>
<UL>
<LI id="s0"><img src="image/ftv2folderopen.gif" align="absmiddle"/>資料上稿</LI>
<xsl:for-each select="CatTree">
	<xsl:apply-templates select="."/>
</xsl:for-each>
</UL>

</body>
</html>
</xsl:template>

<xsl:template match="CatTree">
<LI><xsl:attribute name="id">st<xsl:value-of select="@id" /></xsl:attribute>
		<xsl:value-of disable-output-escaping="yes" select="hyweb:treeImg0(.)" />
	<a href="blank.htm"><xsl:value-of select="@name" /></a>
				<UL>
				<xsl:for-each select="CatTreeNode">
					<xsl:apply-templates select="."/>
				</xsl:for-each>
				</UL>
</LI>
</xsl:template>

	<xsl:template match="CatTreeNode">
		<LI style="white-space:nowrap"><xsl:attribute name="id">s<xsl:value-of select="@id" /></xsl:attribute>
		<xsl:value-of disable-output-escaping="yes" select="hyweb:treeImg(.)" />
			<a>
				<xsl:attribute name="href">ctXMLin.asp?ItemID=<xsl:value-of select="ancestor::CatTree/@id" />&amp;CtNodeID=<xsl:value-of select="@id" /></xsl:attribute>
						<xsl:attribute name="target">ForumToc</xsl:attribute>
					<xsl:value-of select="@name" />
			</a>
			<xsl:if test="@CtNodeKind='C'">
				<UL>
				<xsl:for-each select="CatTreeNode">
					<xsl:apply-templates select="."/>
				</xsl:for-each>
				</UL>
			</xsl:if>
		</LI>
	</xsl:template>

 
<msxsl:script language="VBScript" implements-prefix="hyweb"><![CDATA[
Function treeImg(nodeList) 
  Set myNode=nodeList(0)
   leadStr=myNode.attributes.getNamedItem("leadStr").value
   rStr = ""
   for i = 1 to len(leadStr)-1
   	rStr = rStr & "<img src=""image/t" & mid(leadStr,i,1) & ".gif"" align=""absmiddle""/>"
   next
   if myNode.attributes.getNamedItem("CtNodeKind").value="C" then
	rStr = rStr & "<img id=""pms" & myNode.attributes.getNamedItem("id").value _
		& """ src=""image/fopen" & right(leadStr,1) & ".gif"" align=""absmiddle""/>"
 	rStr = rStr & "<img src=""image/ftv2folderclosed.gif"" align=""absmiddle""/>"
  else
	rStr = rStr & "<img id=""pms" & myNode.attributes.getNamedItem("id").value _
		& """ src=""image/t" & right(leadStr,1) & ".gif"" align=""absmiddle""/>"
 	rStr = rStr & "<img src=""image/openfold.gif"" align=""absmiddle""/>"
   end if
   treeImg = rStr
End Function

Function treeImg0(nodeList) 
  Set myNode=nodeList(0)
   leadStr=myNode.attributes.getNamedItem("leadStr").value
   rStr = ""
   for i = 1 to len(leadStr)-1
   	rStr = rStr & "<img src=""image/t" & mid(leadStr,i,1) & ".gif"" align=""absmiddle""/>"
   next
   if myNode.attributes.getNamedItem("CtNodeKind").value="C" then
	rStr = rStr & "<img id=""pmst" & myNode.attributes.getNamedItem("id").value _
		& """ src=""image/fopen" & right(leadStr,1) & ".gif"" align=""absmiddle""/>"
 	rStr = rStr & "<img src=""image/ftv2folderclosed.gif"" align=""absmiddle""/>"
  else
	rStr = rStr & "<img id=""pmst" & myNode.attributes.getNamedItem("id").value _
		& """ src=""image/t" & right(leadStr,1) & ".gif"" align=""absmiddle""/>"
 	rStr = rStr & "<img src=""image/openfold.gif"" align=""absmiddle""/>"
   end if
   treeImg0 = rStr
End Function

]]></msxsl:script>    
</xsl:stylesheet>


