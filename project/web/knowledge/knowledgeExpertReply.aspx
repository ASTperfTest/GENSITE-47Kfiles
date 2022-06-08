<%@ Page Language="VB" AutoEventWireup="false" CodeFile="knowledgeExpertReply.aspx.vb" Inherits="knowledge_knowledgeExpertReply" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>農業知識入口網</title>
    <link href="/xslGip/style3/css/mailQ.css" rel="stylesheet" type="text/css" />
</head>
<body>
  <div class="Question">
    <div class="head">
      <h1><img src="/xslGip/style3/images/mailq_logo.gif" alt="農業知識入口網" /></h1>
      <h2><img src="/xslGip/style3/images/mailq_h1.gif" alt="農業知識家 專家補充" /></h2>
    </div>

    <form id="form1" runat="server">
      <table class="body" summary="提問內容">
      <tr>
        <th scope="row">專家</th>
        <td><asp:Label ID="ExpertName" runat="server" Text=""></asp:Label></td>
      </tr>
      <tr>
        <th scope="row">問題標題</th>
        <td><asp:Label ID="QuestionTitle" runat="server" Text=""></asp:Label></td>
      </tr>
      <tr>
        <th scope="row">問題內容</th>
        <td>
          <asp:Label ID="QuestionContent" runat="server" Text=""></asp:Label>
        </td>
      </tr>
      <tr>
        <th scope="row">原問題連結</th>
        <td><asp:HyperLink ID="QuestionLink" runat="server" Target="_blank"></asp:HyperLink></td>
      </tr>
      <tr>
        <th scope="row">專家補充</th>
        <td>
          <asp:TextBox ID="ExpertAddText" runat="server" Columns="30" Rows="10" Wrap="true" TextMode="MultiLine"></asp:TextBox>     
          <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="專家補充為必填欄位" ControlToValidate="ExpertAddText"  />
        </td>
      </tr>
      </table>
      <asp:ImageButton ID="SubmitBtn" runat="server" ImageUrl="/xslGip/style3/images/btn_submit.gif" AlternateText="確定送出" />
      <asp:ImageButton ID="ResetBtn" runat="server" ImageUrl="/xslGip/style3/images/btn_reset.gif" AlternateText="重新輸入" />
    </form>
  </div> 
</body>
</html>
