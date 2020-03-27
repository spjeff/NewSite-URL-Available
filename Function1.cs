/* from
 * 
 * https://medium.com/ng-sp/serverless-azure-function-and-sharepoint-csom-7c176b63cfc9
 * https://www.anupams.net/app-only-policy-with-tenant-level-permissions-in-sharepoint-online/
 * https://www.c-sharpcorner.com/article/connect-to-sharepoint-online-site-with-app-only-authentication/
 * https://piyushksingh.com/2018/12/26/register-app-in-sharepoint/
 * https://www.eliostruyf.com/using-the-latest-sharepoint-pnp-core-online-dependency-in-your-azure-functions/
 * 
 */


using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

// CSOM
using Microsoft.SharePoint.Client;

// PNP
using OfficeDevPnP.Core;

namespace NewSite_URL_Available
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {

            // O365 Tenant
            var appid = "6f774955-1a5a-4639-b07e-d7df9cc16f5b";
            var appsecret = "SQOgbwKp4bk5hhOLJ8IMD7N2cQJzpQBdpGFiHWIwELk=";


            // SPO Target
            string name = null;
            string siteUrl = "https://spjeff.sharepoint.com/";
            var am = new AuthenticationManager();
            var context = am.GetAppOnlyAuthenticatedContext(siteUrl, appid, appsecret);
            //using ()
            //{
                // Open SharePoint client web
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();
                name = web.Id.ToString();
            //}


            // Reply
            log.LogInformation("C# HTTP trigger function processed a request.");
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            return name != null
                ? (ActionResult)new OkObjectResult($"Hello, {name}")
                : new BadRequestObjectResult("Please pass a name on the query string or in the request body");
                
        }
    }
}

