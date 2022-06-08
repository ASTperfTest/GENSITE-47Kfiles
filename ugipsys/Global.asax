<%@ Application Language="C#" %>
<%@ Import Namespace="System.Web.Security" %>
<%@ Import Namespace="System.Security.Principal" %>
<%@ Import Namespace="System.Runtime.Serialization" %>
<%@ Import Namespace="System.Runtime.Serialization.Formatters.Binary" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Web.Configuration" %>


<script runat="server">
    
    void Application_Start(object sender, EventArgs e) 
    {
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();        
    }    
    
</script>
