<%@ Page Language="C#" AutoEventWireup="true" CodeFile="historylog.aspx.cs" Inherits="kmactivity_history_historylog" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>未命名頁面</title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="/js/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" >
        function GoTop() {
            window.location.href="/kmactivity/history/activityrankdetail.aspx";
        }
    
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <table width="100%"><tr><td class ="content" width="100%"><br/>
    <div class="pantology2">
    <div class="Page">
    <input type="button" value="回上一頁" onclick="javascript:history.go(-1)"/>&nbsp;
    <input type="button" value="回排行榜" onclick="GoTop()" /> <br />
     帳號：<asp:label ID="loginid" runat="server" text=""></asp:label> &nbsp;&nbsp; 
    姓名/暱稱：<asp:label ID="userName" runat="server" text=""></asp:label>&nbsp;&nbsp; 
    E-mail：<asp:label ID="userMail" runat="server" text=""></asp:label> <br /><br />
    <a href="/kmactivity/history/historylogexport.aspx?type=all&accountid=<%=account_id %>" >匯出報表</a> <br />
    <asp:label ID="treasureTable" runat="server" text=""></asp:label>
     
            
      <hr />   
      <div>
      第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
        <asp:HyperLink ID="PreviousLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
          <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
        </asp:HyperLink>
        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
        <asp:HyperLink ID="NextLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
          <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
        </asp:HyperLink>，每頁                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
          <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
          <asp:ListItem Value="30">30</asp:ListItem>
          <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>筆資料
        </div>
        <asp:Label ID="TableText" runat="server" Text="" />
      </br>
      <div class="top">
        <a href="#" title="top">top</a>
      </div>   
    </div>
    <input type="hidden" value="" id="ReverseQuery" name="ReverseQuery" />   
    <input type="hidden" value="false" id="ForQuery" name="ForQuery" /> 
    <input type="hidden" value="false" id="ForChange" name="ForChange" />   
    </div>
    </td></tr></table>
    </form>
</body>
</html>
