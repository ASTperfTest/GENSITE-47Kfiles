<%@ Application Language="C#" %>
<%@ Import Namespace="System.Web.Security" %>
<%@ Import Namespace="System.Security.Principal" %>
<%@ Import Namespace="System.Runtime.Serialization" %>
<%@ Import Namespace="System.Runtime.Serialization.Formatters.Binary" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="log4net" %>
<%@ Import Namespace="System.Web.Configuration" %>

<script runat="server">
    
    void Application_Start(object sender, EventArgs e) 
    {
        // 應用程式啟動時執行的程式碼
        log4net.Config.XmlConfigurator.Configure();
    }
        
    void Application_Error(object sender, EventArgs e) 
    { 
        Exception ex = Server.GetLastError();
		ILog log4 = LogManager.GetLogger("Dr");
		log4.Error(ex.ToString());
                  
    }

</script>
