<%@ Page Title="" Language="C#" MasterPageFile="Default.master" AutoEventWireup="true" CodeFile="result.aspx.cs" Inherits="result" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">
    <div id="content">
    </div>

    <script type="text/javascript">
        $(document).ready(function() {
            //Tab Menu
            $("#tabInfo").addClass("btnDistrict");
            $("#tabList").addClass("btnDistrict");
            $("#tabVote").addClass("btnDistrict");
            $("#tabFinal").addClass("btnDistrict now");
        });
    </script>

</asp:Content>
