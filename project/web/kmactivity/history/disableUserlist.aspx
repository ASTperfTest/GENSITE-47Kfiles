<%@ Page Language="C#" AutoEventWireup="true" CodeFile="disableUserlist.aspx.cs" Inherits="kmactivity_history_disableUserlist" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>未命名頁面</title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="/js/jquery-1.3.2.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <table width="100%"><tr><td class ="content" width="100%"><br/>
    <div class="pantology2">
    <div class="Page">
    <input type="button" value="回上一頁" onclick="javascript:history.go(-1)"/>&nbsp;
    <input type="button" value="回排行榜" onclick="location.href='/kmactivity/history/activityrankdetail.aspx" /> <br />

    <asp:label ID="treasureTable" runat="server" text=""></asp:label>
     
            
      <hr />   
      <div>
      
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
    </form>
</body>
</html>
