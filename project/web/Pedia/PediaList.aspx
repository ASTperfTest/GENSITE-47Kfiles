<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PediaList.aspx.vb" Inherits="Pedia_PediaList" Title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
    <div class="path">
        目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>
    <div class="pantology">
        <div class="head"></div>
        <div class="body">
            <h2>~小百科規則~</h2>
            <div class="sublayout">
                <img src="Icon.png">
                <h4>農業知識小百科功能公告</h4>
                <p> 將農業知識庫和農業主題館中看到的「農業詞彙」推薦給小百科，例如提出對「摘心」、「組識培養」有疑惑，可使用「推薦詞彙」發表詢問，而網站的會員可使用「我來補充」的編輯行列，到小百科裡一一來解釋農業詞彙的意義。...<a href="PediaInfoDetail.aspx">詳全文</a></p>
            </div>
        </div>
        <div class="body">
            <h2>
                農業知識小百科</h2>
            <div class="browseby">
                <!--
        <label for="select2">瀏覽條件篩選 依筆畫：</label>
        <select name="select2" id="select2">
          <option selected>請選擇首字筆劃數</option>
          <option>1-3劃</option>
          <option>4-6劃</option>
        </select>
  -->
                <table width="100%">
                    <tr>
                        <td align="left">
                            <label for="select2">
                                開放補充：</label>
                            <asp:dropdownlist runat="server" id="isOpenDDL" autopostback="true">
          <asp:ListItem Value="" Text="請選擇" />
          <asp:ListItem Value="Y" Text="是" />
          <asp:ListItem Value="N" Text="否" />
        </asp:dropdownlist>
                        </td>
                        <td>
                            <ul class="Function2">
                                <li><a href="/Pedia/PediaQuery.aspx" class="retrieve">詞目查詢</a></li>
                            </ul>
                        </td>
                    </tr>
                </table>
            </div>
            <!--分頁/開始 -->
            <div class="Page">
                第<asp:label id="PageNumberText" runat="server" cssclass="Number" />/<asp:label id="TotalPageText" runat="server" cssclass="Number" />頁， 共<asp:label id="TotalRecordText" runat="server" cssclass="Number" />筆資料，
                <asp:hyperlink id="PreviousLink" runat="server">
          <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
          <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
        </asp:hyperlink>
                跳至第<asp:dropdownlist id="PageNumberDDL" autopostback="true" runat="server" />
                頁 &nbsp;
                <asp:hyperlink id="NextLink" runat="server">
          <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
          <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
        </asp:hyperlink>
                ，每頁
                <asp:dropdownlist id="PageSizeDDL" autopostback="true" runat="server">
          <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
          <asp:ListItem Value="30">30</asp:ListItem>
          <asp:ListItem Value="50">50</asp:ListItem>
        </asp:dropdownlist>
                筆資料
            </div>
            <asp:label id="TableText" runat="server" text="" />
        </div>
        <div class="foot">
        </div>
    </div>
</asp:Content>
