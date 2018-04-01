using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Web.Services.Protocols;

namespace SPOPermissionReport
{


    public class WSHelper
    {

        public static WSType CreateWebService<WSType>(Uri siteUrl) where WSType : SoapHttpClientProtocol, new()
        {
            WSType webService = new WSType();
            webService.Credentials = CredentialCache.DefaultNetworkCredentials;

            string webServiceName = typeof(WSType).Name;
            webService.Url = string.Format("{0}/ _vti_bin /{ 1}.asmx", siteUrl, webServiceName);

            return webService;
        }

        public static void TestWS()
        {
            ServiceReference1.UserProfileServiceSoapClient c = new ServiceReference1.UserProfileServiceSoapClient();
            //c.Url
        }
    }

}