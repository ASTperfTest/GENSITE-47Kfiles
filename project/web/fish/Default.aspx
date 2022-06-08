<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>豆仔水族箱</title>
	<script type="text/javascript" src="http://kmweb.coa.gov.tw/ViewCounter.aspx"></script>
    <script src="js/swfobject/swfobject.js" type="text/javascript"></script>
	<script type="text/javascript">
var gaJsHost=(("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
		    var pageTracker = _gat._getTracker("UA-9195501-1");
		    pageTracker._trackPageview();
        }
        catch (err) {alert(err.description ); }   
</script>
    <link href="css/reset.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="#ffffff">
    <center>
        <form id="form1" runat="server">
        <div>
            <div id="main">
                Game</div>
            <asp:Literal ID="GameStr" runat="server"></asp:Literal>
        </div>
    </center>
    </form>
</body>
</html>
