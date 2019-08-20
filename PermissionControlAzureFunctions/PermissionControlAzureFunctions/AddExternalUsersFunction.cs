using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using OfficeDevPnP.Core.Utilities;

namespace PermissionControlAzureFunctions
{
    public static class AddExternalUsersFunction
    {
        [FunctionName("AddExternalUsers")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "havePermission")]HttpRequestMessage req, TraceWriter log)
        {
            try
            {
                // Gets data from request body.
                dynamic data = await req.Content.ReadAsAsync<object>();

                string siteUrl = data.SiteUrl;
                string currentEmail = data.CurrentUser_EmailAddress;
                if (String.IsNullOrEmpty(siteUrl) || String.IsNullOrEmpty(currentEmail))
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Please pass parametes site URL and Email Address in request body!");

                // Fetches client id and client secret from app settings.
                string clientId = Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
                string clientSecret = Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);
                string urlAdminSite = Environment.GetEnvironmentVariable("UrlAdminSite", EnvironmentVariableTarget.Process);

                // Obtains client context using the client id and client secret.
                var ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(urlAdminSite, clientId, clientSecret);

                Tenant tenant = new Tenant(ctx);
                SiteProperties siteProp = tenant.GetSitePropertiesByUrl(siteUrl, true);
                Site site = tenant.GetSiteByUrl(siteUrl);
                Web web = site.RootWeb;
                RoleAssignmentCollection roleAssignments = web.RoleAssignments;

                //Get the current user by email address.
                User currentUser = web.EnsureUser(currentEmail);

                ctx.Load(site, s => s.RootWeb, s => s.RootWeb.CurrentUser);
                ctx.Load(siteProp);
                ctx.Load(currentUser);
                ctx.Load(web, w => w.CurrentUser, w => w.SiteGroups, w => w.Url, w => w.Title);
                ctx.Load(roleAssignments, roleAssignement => roleAssignement.Include(r => r.Member, r => r.RoleDefinitionBindings));
                ctx.ExecuteQuery();

                var allowedListDomainFromSite = siteProp.SharingAllowedDomainList.Split(',').Select(x => x.Trim().ToUpper()).ToList();

                if (!siteProp.SharingCapability.ToString().Equals("Disabled"))
                {

                    if (allowedListDomainFromSite.Count() != 0)
                    {
                        if (currentUser.IsSiteAdmin || CheckUserInAdminGroups(ctx, web, roleAssignments, currentUser)) // check if the current user have full control
                        {
                            return req.CreateResponse(HttpStatusCode.OK, true);
                        }
                    }

                }
                else return req.CreateResponse(HttpStatusCode.OK, false);

                return req.CreateResponse(HttpStatusCode.OK, false);
            }
            catch (Exception e)
            {
                return req.CreateResponse(HttpStatusCode.InternalServerError, e.Message);
            }
        }

        private static bool CheckUserInAdminGroups(ClientContext ctx, Web web, RoleAssignmentCollection roleAssignments, User currentUser)
        {
            if (roleAssignments != null && roleAssignments.Count != 0)
            {
                foreach (RoleAssignment ra in roleAssignments)
                {
                    //Load users foreach Group - Group Name: ra.Member.Title
                    UserCollection usersCollection = web.SiteGroups.GetByName(ra.Member.Title).Users;
                    ctx.Load(usersCollection);
                    ctx.ExecuteQuery();
                    //Check Permission Groups - rd.Name: Permission Level
                    foreach (RoleDefinition rd in ra.RoleDefinitionBindings)
                    {
                        if (rd.Name.Equals("Full Control") && (usersCollection.GetByEmail(currentUser.Email) != null))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
    }
}
