<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master"
 AutoEventWireup="true" CodeFile="Picture_Detail.aspx.cs" Inherits="Century_Picture_Detail" 
  EnableViewState="false"
 %>

<%@ Register Src="TabText.ascx" TagName="TabMenu" TagPrefix="uc1" %>
<%@ Register Src="path.ascx" TagName="path" TagPrefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
<script language="javascript" type="text/javascript" src="/js/yoxviewPic/jquery.yoxview-2.2.min.js"></script>
<script language="javascript" type="text/javascript" src="/js/yoxviewPic/yoxview-init.js"></script>
<script src="/nocopy.js" language="javascript"></script>
<link rel="Stylesheet" href="/js/yoxviewPic/yoxview.css" />
<link rel="stylesheet" type="text/css" href="css/i_album.css" />
<script language="javascript" type="text/javascript">
    $(document).ready(function() {
        $(".yoxview").yoxview({
            videoSize: { maxwidth: 720, maxheight: 560 },
            autoHideInfo: false,
            autoHideMenu: false,
            lang: 'zh-tw'
        });
        $(".yoxview").show();
    });    (jQuery)
    </script>
    <!--中間內容區-->
    <!--  目前位置  -->
    <uc2:path runat="server" ID="path1" />
    <!--標籤區-->
    <uc1:TabMenu runat="server" id="TabMenu1" />
    <div>
        <asp:Label runat="server" ID="labTAGs" CssClass="category"></asp:Label>
    </div>
    <div class="yoxview" style="display:none;">
        <asp:Literal runat="server" ID="litPictures"></asp:Literal>
    </div>
    &nbsp;
    <div style="clear:both"><br/>
    <p>(以上資料皆由豐年社提供)</p>
    </div>
</asp:Content>