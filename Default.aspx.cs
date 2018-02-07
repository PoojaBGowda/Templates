using Microsoft.SharePoint.Client;
using System;
using System.Net;
using System.Text;
using System.Security;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Configuration;

namespace TemplatesSelectionWeb
{
    public partial class Default : System.Web.UI.Page
    {
        public string tenantStr = "";
        public string webUrl = "";
        public string tenantAdminUri = "";
        string hostWebUrl = "";
        public string userName = "";
        public string userPassword = "";

        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                templatedropdown.Items.Add(new System.Web.UI.WebControls.ListItem("Cintranet", "Cintranet"));
                templatedropdown.Items.Add(new System.Web.UI.WebControls.ListItem("Eldon", "Eldon"));
                templatedropdown.Items.Add(new System.Web.UI.WebControls.ListItem("Gasum", "Gasum"));
                templatedropdown.Items.Add(new System.Web.UI.WebControls.ListItem("Hologic", "Hologic"));
                templatedropdown.Items.Add(new System.Web.UI.WebControls.ListItem("Northern Woods", "NorthernWoods"));
                templatedropdown.Items.Add(new System.Web.UI.WebControls.ListItem("NorthLight", "NorthLight"));
                templatedropdown.Items.Add(new System.Web.UI.WebControls.ListItem("Woodland", "Woodland"));
                urldropdown.Items.Add(new System.Web.UI.WebControls.ListItem("sites", "sites"));
                urldropdown.Items.Add(new System.Web.UI.WebControls.ListItem("teams", "teams"));
            }
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
            using (ClientContext ctx = spContext.CreateUserClientContextForSPHost())
            {
                hostWebUrl = Page.Request["SPHostUrl"];
                tenantStr = hostWebUrl.ToLower().Replace("-my", "").Substring(8);
                tenantStr = tenantStr.Substring(0, tenantStr.IndexOf("."));
                webUrl = String.Format("https://{0}.sharepoint.com/{1}/{2}", tenantStr, urldropdown.SelectedValue, urltxt.Text);
                tenantAdminUri = String.Format("https://{0}-admin.sharepoint.com", tenantStr);
                urlbl.Text = String.Format("https://{0}.sharepoint.com/", tenantStr);
                prmownertxt.Text = ConfigurationManager.AppSettings["Email"];
                prmownertxt.Enabled = false;
            }

        }
        public Boolean check_site()
        {
            using (ClientContext tenantContext = new ClientContext(tenantAdminUri))
            {
                SecureString password = new SecureString();
                foreach (char c in userPassword.ToCharArray())
                    password.AppendChar(c);
                tenantContext.Credentials = new SharePointOnlineCredentials(userName, password);
                var tenant = new Tenant(tenantContext);
                SPOSitePropertiesEnumerable sitePropEnumerable = tenant.GetSiteProperties(0, true);
                tenantContext.Load(sitePropEnumerable);
                tenantContext.ExecuteQuery();
                foreach (SiteProperties property in sitePropEnumerable)
                {
                    if (property.Url.Equals(webUrl))
                        return true;

                }
                return false;
            }
        }

        protected void Reset_Click(object sender, EventArgs e)
        {
            sitenmetxt.Text = string.Empty;
            urltxt.Text = string.Empty;
            descr.Text = string.Empty;
        }

        protected void Create_Click(object sender, EventArgs e)
        {
            userName = System.Configuration.ConfigurationManager.AppSettings["Email"]; //userName 
            userPassword = System.Configuration.ConfigurationManager.AppSettings["Password"]; ; //userPWD
            bool a = check_site();
            if (a == false)
            {

                //change webJobName to your WebJob name 
                webUrl = String.Format("https://{0}.sharepoint.com/{1}/{2}", tenantStr, urldropdown.SelectedValue, urltxt.Text);

                string webJobName = templatedropdown.SelectedValue;

                //Change this URL to your WebApp hosting the  
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("https://businesstemplate.scm.azurewebsites.net/api/triggeredwebjobs/" + webJobName + "/run?arguments=" + System.Web.HttpUtility.UrlEncode(webUrl + " " + tenantAdminUri + " " + userName + " " + userPassword + " " + sitenmetxt.Text + " " + descr.Text));
                request.Method = "POST";
                var byteArray = Encoding.ASCII.GetBytes("$BusinessTemplate:w2duTHEtMYaa2yvQmpjNF0LeE2oi2m4rPzZrRT8mEEErqCT1j33eoFj8eJxr"); //we could find user name and password in Azure web app publish profile
                request.Headers.Add("Authorization", "Basic " + Convert.ToBase64String(byteArray));
                request.ContentLength = 0;
                try
                {
                    var response = (HttpWebResponse)request.GetResponse();
                    errlbl.Text = "Creating " + sitenmetxt.Text + " site....\n you will be notified with a mail after completion.";


                }
                catch (Exception ex)
                {
                    errlbl.Text = ex.Message;
                }
            }
            else
                errlbl.Text = "Site already exists... Try with different name.";

        }

        protected void Cancel_Click(object sender, EventArgs e)
        {
            Response.Redirect(hostWebUrl + "/_layouts/15/viewlsts.aspx?view=14");

        }
    }
}