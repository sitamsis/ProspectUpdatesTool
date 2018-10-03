using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Net;
using System.Configuration;

namespace WTFiserv.Crm.PU.DataAccess
{
    public class Connection
    {

        public OrganizationServiceProxy ConnectToCRM(log4net.ILog log)
        {
            try
            {
                var organizationUri = new Uri(GetURL());

                var credentials = new ClientCredentials();
                credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
                log.Info("Connecting to CRM Organization " + organizationUri.ToString());
                OrganizationServiceProxy _service = new OrganizationServiceProxy(organizationUri, null, credentials, null);
                return _service;
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                return null;
            }
        }

        private String GetURL()
        {
            //URL values
            String server = ConfigurationManager.AppSettings["server"].ToString();
            if (server.Contains("https"))
            {
                //secure port
                server += ":443";
            }
            String organization = ConfigurationManager.AppSettings["organization"].ToString();
            return server + "/" + organization + "/XRMServices/2011/Organization.svc";
        }
    }
}
