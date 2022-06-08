<%@ Page Title="農業知識入口網-好文推薦" Language="C#" MasterPageFile="~/4SideMasterPage.master"
    AutoEventWireup="true" CodeFile="Recommandinfodetail.aspx.cs" Inherits="recommand_Recommandinfodetail" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
    <div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="float: left; width: 80%">
        <ul id="path_menu">
            <li><a href="../mp.asp?mp=1">首頁</a></li>
            <li style='top: 10px;'>></li>
            <li><a href="Recommand_List.aspx">好文推薦</a></li>
        </ul>        
    </div>
        
    <div class="pantology" style="clear:both">
        <div class="head"></div>
        <div class="body">
            <h2><asp:label id="lbtitle" runat="server"></asp:label></h2>
            <div  class="sublayout">
                <div class="Function2">
                    <ul>
                        <li><a href="javascript:history.go(-1);" class="Back" title="回上一頁">回上一頁</a><noscript>本網頁使用SCRIPT編碼方式執行回上一頁的動作，如果您的瀏覽器不支援SCRIPT，請直接使用「Alt+方向鍵向左按鍵」來返回上一頁</noscript>
                        </li>
                    </ul>
                </div>
                <asp:Image id="Image1" runat="server" style="border:#aaa 1px solid"></asp:Image>                
                <h4>&nbsp;</h4>
                <p><asp:Label id="lbxbody" runat="server"></asp:Label></p><br />
            </div>
        </div>
        <div class="foot">
        </div>
    </div>        
    
</asp:Content>