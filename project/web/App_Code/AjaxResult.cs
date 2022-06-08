using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System;


/// <summary>
/// ProcessingResult 的摘要描述
/// </summary>
public sealed class AjaxResult
{

    private bool success = false;
    private string message = "";
    private Object data = null;

    public AjaxResult()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
    }

    public bool Success
    {
        get { return success; }
        set { success = value; }
    }

    public string Message
    {
        get { return message; }
        set { message = value; }
    }

    public Object Data
    {
        get { return data; }
        set { data = value; }
    }
}
