<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false"
    CodeFile="ArticleSearch.aspx.vb" Inherits="ArticleSearch" Title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <link rel="stylesheet" type="text/css" href="/css/jquery.autocomplete.css" />
    <script type="text/javascript" src="/js/jquery.autocomplete.js"></script>
     <script type="text/javascript" language="javascript">
         $(function() {
            $("#ctl00_ContentPlaceHolder1_txtSubject").autocomplete('AutoComplete.aspx', { delay: 10 });
         });
            function IsDate(obj) {
                var Vdate = obj.value;
                if (Vdate == "") return true;
                var MinYear = 1900;
                var Year;
                var Month;
                var Day;
                var reg = /^(\d{4})([-\/])(\d{1,2})\2(\d{1,2})$/;
                if (!reg.test(Vdate)) {
                    alert("日期格式不正確。");
                    obj.focus();
                    return false;
                }
                else {
                    Year = RegExp.$1;
                    Month = RegExp.$3;
                    Day = RegExp.$4;
                    if (Year < MinYear) {
                        alert("日期中年份太小。");
                        obj.focus();
                        return false;
                    }
                    if (!(Month < 13 && Month > 0)) {
                        alert("日期中月份不正確。");
                        obj.focus();
                        return false;
                    }
                    if (!(Day > 0 && Day < 32)) {
                        alert("日期不正確。");
                        obj.focus();
                        return false;
                    }
                    if (Month == 2) {
                        if (0 == Year % 4 && ((Year % 100 != 0) || (Year % 400 == 0))) {
                            if (Day > 29) {
                                alert("2月份不能大於29天。");
                                obj.focus();
                                return false;
                            }
                        }
                        else {
                            if (Day > 28) {
                                alert("2月份不能大於28天。");
                                obj.focus();
                                return false;
                            }
                        }
                    }
                    else {
                        if (Month == 1 || Month == 3 || Month == 5 || Month == 7 || Month == 8 || Month == 10 || Month == 12) {
                            if (Day > 31) {
                                alert("日期不能超過31天。");
                                obj.focus();
                                return false;
                            }
                        }
                        else {
                            if (Day > 30) {
                                alert("日期不能超過30天。");
                                obj.focus();
                                return false;
                            }
                        }
                    }

                }
                return true;
            }             
    </script>   
    <link rel="stylesheet" type="text/css" href="/css/datepicker.css" />
    <script type="text/javascript" src="/js/datepicker.js"></script>
    <script type="text/javascript" language="javascript">
        $(document).ready(function(e) {
            $("#ctl00_ContentPlaceHolder1_txtFrom").datepicker();
            $("#ctl00_ContentPlaceHolder1_txtTo").datepicker();
        })
    </script>
    
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="主題館" href="/SubjectList.aspx">主題館</a>&gt;<a href="#">文章查詢</a></div>
    <h3>
        主題館-文章查詢</h3>
    <table width="100%">
        <tr><td align="right"><asp:HyperLink ID="HyperLink1" runat="server" 
                Font-Size="Small" NavigateUrl="~/SubjectList.aspx">回上一頁</asp:HyperLink></td></tr>
    </table>    
    
    <table>
        <tr><td>
            <asp:Label ID="Label1" runat="server" Text="指定主題館"></asp:Label>
            </td><td>
                <asp:DropDownList ID="ddlKind" runat="server">
                </asp:DropDownList>
                <asp:TextBox ID="txtSubject" runat="server" AutoCompleteType="Disabled"></asp:TextBox>
                <asp:DropDownList ID="ddlSubject" runat="server" visible="false">
                </asp:DropDownList>
            </td></tr>
        <tr><td>
            <asp:Label ID="Label2" runat="server" Text="文章標題"></asp:Label>
            </td><td>
                <asp:TextBox ID="txtTitle" runat="server" Width="350px"></asp:TextBox>
            </td></tr>
        <tr><td>
            <asp:Label ID="Label3" runat="server" Text="文章內文"></asp:Label>
            </td><td>
                <asp:TextBox ID="txtContent" runat="server" Width="350px"></asp:TextBox>
            </td></tr>
        <tr><td>
            <asp:Label ID="Label4" runat="server" Text="張貼日期"></asp:Label>
            </td><td>
                <asp:TextBox ID="txtFrom" runat="server"></asp:TextBox>~
                <asp:TextBox ID="txtTo" runat="server"></asp:TextBox>
            </td></tr>
        <tr><td></td><td>
            <asp:Button ID="btnSearch" runat="server" Text="查詢" OnClientClick="if (!(IsDate(ctl00_ContentPlaceHolder1_txtFrom) && IsDate(ctl00_ContentPlaceHolder1_txtTo))) {return false;}" />
            <asp:Button ID="btnReset" runat="server" Text="重設" />
            </td></tr>
    </table>
</asp:Content>
