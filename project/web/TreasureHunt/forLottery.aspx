<%@ Page Language="C#" AutoEventWireup="true" CodeFile="forLottery.aspx.cs" Inherits="TreasureHunt_forLottery" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <title>未命名頁面</title>
    <style type="text/css"> 
    .present_title{
	color:#F30;
	font-weight:bold;
	font-size:13px;
	white-space:nowrap;
    }
    .chance{
	font-weight:bold;
	color:#F60;
	font-size:15px;
	text-decoration:underline;
	margin:0 3px;
    }
  </style> 
</head>
<body>
    <form id="form1" runat="server">
    <table width="100%"><tr><td class ="content" width="100%"><br/>
    <div class="pantology2">
    <div class="Page">
    <input type="button" value="回上一頁" onclick="javascript:history.go(-1)"/>&nbsp;
    <input type="button" value="回排行榜" onclick="location.href='/treasureHunt/activityrankdetail.aspx?avtivityid=<%=avtivityId %>'" /> &nbsp;
    <input type="button" id="lotterytrans" value="回獎品兌換頁" onclick="window.location ='/treasurehunt/lotterydetail.aspx?avtivityid=<%=avtivityId %>'" />
    <br />
    帳號：<asp:label ID="loginid" runat="server" text=""></asp:label> &nbsp;&nbsp; 
    姓名/暱稱：<asp:label ID="userName" runat="server" text=""></asp:label>&nbsp;&nbsp; 
    E-mail：<asp:label ID="userMail" runat="server" text=""></asp:label> 
	<br/><h4><asp:label ID="suit" runat="server" text=""></asp:label>&nbsp;&nbsp;&nbsp;&nbsp;已兌換套數:<asp:label ID="ForLottery" runat="server" text=""></asp:label></h4>
    <asp:label ID="LabelGiftVotes" runat="server" text=""></asp:label>
     

    </div>
    </div>
    </td></tr></table>
    </form>
</body>
</html>
