<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MIP.aspx.cs" Inherits="MIP" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農業知識入口網 －小知識串成的大力量</title>
</head>
<body>
    <form id="form1" runat="server" name="name_form1">
    <div>
        <table id="tb1">
            <tr>
                <td>
                    <img alt="FBIcon" title="FBIcon" src="../images/FBIcon.png" />
                </td>
                <td>
                    <p>
                        朋友們，介紹你們一個好站，可以讓你輕鬆習得農業知識，掌握最新農業資訊，快來加入農業知識入口網吧～～
                    </p>
                </td>
            </tr>
        </table>
    </div>
    </form>
    <script type="text/javascript">
        var joinMemberUrl = '<%=joinMemberUrl %>'
        if (joinMemberUrl == '') {
        }
        else {
            document.getElementById('tb1').style.display = 'none';
            window.location.href = joinMemberUrl;
        }
    </script>
</body>
</html>
