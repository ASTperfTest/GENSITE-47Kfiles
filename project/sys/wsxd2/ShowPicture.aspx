<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ShowPicture.aspx.cs" Inherits="ShowPicture" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>瀏覽圖片</title>
    <style type="text/css">

<!–[if IE]>
       <!–[if lte IE 6]>
         .AdjustSize
         {
            width:80px;
           
         }
         .ShowSize
         {
            width:640px;
         }         
       <![endif]–>
<![endif]–>        
        
 .Up
 {
    padding-bottom:135px;
 }
 .AdjustSize
 {
    max-width:80px;
	max-height:80px;
   
 }
 .ShowSize
 {
    max-width:640px;
    max-height:640px;
 }
 
 </style>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <div style="text-align: center">
            <ajaxToolkit:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender1" runat="server"
                TargetControlID="UpdatePanelPhoto">
                <Animations>                             
                        <OnUpdating>
                            <Sequence>
                                <StyleAction Attribute="overflow" Value="hidden" />
                                <Parallel duration=".2" Fps="10">
                                        <FadeOut AnimationTarget="photo" minimumOpacity=".2" />                                                                 
                                </Parallel>
                            </Sequence>
                        </OnUpdating> 
                        <OnUpdated>
                            <Sequence>
                                <Parallel duration=".2" Fps="10">                               
                                        <FadeIn AnimationTarget="photo" minimumOpacity=".2" />                                                               
                                </Parallel>  
                            </Sequence>
                        </OnUpdated>                    
                </Animations>
            </ajaxToolkit:UpdatePanelAnimationExtender>
            <asp:Panel ID="PanelImage" runat="server" HorizontalAlign="Center">
                <br />
                <br />
                <div id="photo">
                    <asp:UpdatePanel ID="UpdatePanelPhoto" runat="server">
                        <ContentTemplate>
                            <table width="640px">
                                <tr align="right">
                                    <td>
                                        第
                                        <asp:DropDownList ID="ddlPicture" runat="server" OnSelectedIndexChanged="DDLPicture_SelectedIndexChanged"
                                            AutoPostBack="True">
                                        </asp:DropDownList>
                                        張 (共
                                        <asp:Label ID="LabelPicTotal" runat="server"></asp:Label>
                                        張)
                                    </td>
                                </tr>
                                <tr align="center">
                                    <td>
                                        <asp:Image ID="ImageShow" runat="server" CssClass="ShowSize"/><br />
                                    </td>
                                </tr>
                                <tr align="center">
                                    <td>
                                        <asp:Label ID="TitleShow" runat="server"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </ContentTemplate>
                        <Triggers>
                            <asp:AsyncPostBackTrigger ControlID="ImageDataList" EventName="SelectedIndexChanged" />
                            <asp:AsyncPostBackTrigger ControlID="PreImageButton" EventName="Click" />
                            <asp:AsyncPostBackTrigger ControlID="NextImageButton" EventName="Click" />
                            <asp:AsyncPostBackTrigger ControlID="ddlPicture" EventName="SelectedIndexChanged" />
                        </Triggers>
                    </asp:UpdatePanel>
                </div>
            </asp:Panel>
            <hr />
            <asp:UpdatePanel ID="UpdatePanelPhotoList" runat="server">
                <ContentTemplate>
                    <asp:Table ID="Table1" runat="server" Height="252px" Width="470px">
                        <asp:TableRow ID="TableRow1" runat="server">
                            <asp:TableCell ID="TableCell1" runat="server">
							<asp:HiddenField ID="HiddenFieldRowNumber" runat="server" Value='' />
                                <asp:ImageButton ID="PreImageButton" runat="server" ImageUrl="../public/preImage.png"
                                    Visible="false" OnClick="PreImageButton_Click" CssClass="Up" /></asp:TableCell>
                            <asp:TableCell ID="TableCell2" runat="server">
                                <asp:DataList ID="ImageDataList" runat="server" RepeatDirection="Horizontal" OnSelectedIndexChanged="ImageDataList_SelectedIndexChanged"
                                    BorderColor="#404040" DataKeyField="rowNum" OnItemDataBound="ImageDataList_ItemDataBound">
                                    <ItemTemplate>
                                        <table style="table-layout: fixed;" >
                                            <tr align="center" valign="top">
                                                    <td valign="middle" style="background-color: Silver; height: 100px; width:100px; background-image: url('../public/background.jpg');">
                                                        <asp:ImageButton ID="ImageButton" runat="server" CommandName="select" CssClass="AdjustSize" ImageUrl='<%# Eval("imgUrl") %>'
                                                            AlternateText='<%# Eval("NFileName") %>' />
                                                    </td>
                                            </tr>
                                            <tr align="center" valign="top" style="height: 200px">
                                                <td valign="top">
                                                    <asp:Label ID="LabelTitle" runat="server" Text='<%# Eval("aTitle") %>' Font-Size="Small"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:HiddenField ID="HiddenFieldFileName" runat="server" Value='<%# Eval("NFileName") %>' />
                                                </td>
                                            </tr>
                                        </table>
                                    </ItemTemplate>
                                </asp:DataList>
                            </asp:TableCell>
                            <asp:TableCell ID="TableCell3" runat="server">
                                <asp:ImageButton ID="NextImageButton" runat="server" ImageUrl="../public/nextImage.png"
                                    Visible="false" OnClick="NextImageButton_Click" CssClass="Up" /></asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </ContentTemplate>
            </asp:UpdatePanel>
            <br />
            <br />
         
        </div>
    </form>
</body>
</html>
