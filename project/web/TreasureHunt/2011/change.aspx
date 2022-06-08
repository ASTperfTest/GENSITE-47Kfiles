<%@ Page Language="C#" AutoEventWireup="true" CodeFile="change.aspx.cs" Inherits="TreasureHunt_2011_change" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body onunload="end()" style="padding:0;">
    <script type="text/javascript" src="/js/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" >
        function end() {
            top.location.href = "myspace.aspx?activityid=3";

//            var oHead = window.parent.document.getElementsByTagName('head').item(0); 
//            var sc= window.parent.document.createElement("script"); 
//            sc.type = "text/javascript";
//            sc.text = "rr =function (){alert(parent.location);}";
//            oHead.appendChild(sc);
//            window.parent.rr();
            

        }
    </script>
    <form id="form1" runat="server">
    <div align="center">
    <embed width="400" height="320" align="middle" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" allowfullscreen="false" allowscriptaccess="sameDomain" wmode="transparent" quality="high" src="<%=flashfile %>" name="banner">
    <br /><br />
    <input type="button" onclick="end();" value="關閉"/>
    </div>
    </form>
</body>
</html>
