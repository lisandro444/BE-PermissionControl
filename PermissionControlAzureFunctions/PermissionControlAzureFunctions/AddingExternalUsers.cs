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
    public static class AddingExternalUsers
    {
        [FunctionName("AddingExternalUsers")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "AddExternalUser")]HttpRequestMessage req, TraceWriter log)
        {
            try
            {
                // Gets data from request body.
                dynamic data = await req.Content.ReadAsAsync<object>();

                string siteUrl = data.SiteUrl;
                string currentEmail = data.CurrentUser_EmailAddress;
                string groupName = data.GroupName;
                if (String.IsNullOrEmpty(siteUrl) || String.IsNullOrEmpty(currentEmail))
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Please pass parametes site URL and Email Address in request body!");

                // Fetches client id and client secret from app settings.
                string clientId = Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
                string clientSecret = Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);
                string urlAdminSite = Environment.GetEnvironmentVariable("UrlAdminSite", EnvironmentVariableTarget.Process);

                // Obtains client context using the client id and client secret.
                var ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(urlAdminSite, clientId, clientSecret);

                Tenant tenant = new Tenant(ctx);
                Site site = tenant.GetSiteByUrl(siteUrl);
                Web web = site.RootWeb;
                Group group = null;
       
                ctx.Load(web, w => w.SiteGroups);
               
                ctx.ExecuteQuery();

                // Check if the group exists
                group = web.SiteGroups.GetByName(groupName);

                // If it doesn't exist, add it
                if (group == null)
                {
                    if(groupName == "SCJ External Contribute")
                    {
                        // Create the Group
                        CreateContributePermissionLevel(ctx);
                        AddGroup(ctx, groupName);
                        ctx.Web.AddUserToGroup(group, "lisandro.rossi@spr.com");
                    }

                    if (groupName == "SCJ External Read")
                    {
                        // Create the Group and Add the User
                        CreateContributePermissionLevel(ctx);
                        AddGroup(ctx, groupName);

                        ctx.Web.AddUserToGroup(group, "lisandro.rossi@spr.com");
                    }

                }
                else // Just Add the user to group
                {
                    ctx.Web.AddUserToGroup(group, "lisandro.rossi@spr.com");
                    // testGroup.InviteExternalUser("mahesh.srinivasan@spr.com");  
                }

                return req.CreateResponse(HttpStatusCode.OK, false);

            }
            catch (Exception e)
            {
                return req.CreateResponse(HttpStatusCode.InternalServerError, e.Message);
            }
        }


        private static bool checkUserDomainFrom(string emailAddress, string SiteUrl)
        {

            // check if the domain from email adress is inclued into list white domain

            return true;

        }

        private static void AddGroup(ClientContext ctx, string groupName)
        {
            //add groupo to a site with the same name as permission level
            ctx.Site.RootWeb.AddGroup(groupName, "Permission Control - Custom Group for External User", false, true, false);
            ctx.Web.AddPermissionLevelToGroup(groupName, "Permission Control - Custom Permission for External User", true);

        }

        private static void CreateContributePermissionLevel(ClientContext ctx)
        {
            // Create New Custom Permission Level
            RoleDefinitionCreationInformation roleDefinitionCreationInformation = new RoleDefinitionCreationInformation();
            BasePermissions perms = new BasePermissions();
            perms.Set(PermissionKind.AddListItems);
            perms.Set(PermissionKind.EditListItems);
            perms.Set(PermissionKind.DeleteListItems);
            perms.Set(PermissionKind.ViewListItems);
            perms.Set(PermissionKind.ApproveItems);
            perms.Set(PermissionKind.OpenItems);
            perms.Set(PermissionKind.ViewVersions);
            perms.Set(PermissionKind.DeleteVersions);
            perms.Set(PermissionKind.CreateAlerts);
            perms.Set(PermissionKind.ViewPages);
            roleDefinitionCreationInformation.BasePermissions = perms;
            roleDefinitionCreationInformation.Name = "SCJ External Contribute";
            roleDefinitionCreationInformation.Description = "Custom Permission Level - SCJ External Contribute";
            ctx.Site.RootWeb.RoleDefinitions.Add(roleDefinitionCreationInformation);

        }

        private static void CreateReadPermissionLevel(ClientContext ctx)
        {
            // Create New Custom Permission Level
            RoleDefinitionCreationInformation roleDefinitionCreationInformation = new RoleDefinitionCreationInformation();
            BasePermissions perms = new BasePermissions();
            perms.Set(PermissionKind.ViewListItems);
            perms.Set(PermissionKind.ApproveItems);
            perms.Set(PermissionKind.OpenItems);
            perms.Set(PermissionKind.ViewVersions);
            perms.Set(PermissionKind.DeleteVersions);
            perms.Set(PermissionKind.CreateAlerts);
            perms.Set(PermissionKind.ViewPages);
            roleDefinitionCreationInformation.BasePermissions = perms;
            roleDefinitionCreationInformation.Name = "SCJ External Read";
            roleDefinitionCreationInformation.Description = "Custom Permission Level - SCJ External Read";
            ctx.Site.RootWeb.RoleDefinitions.Add(roleDefinitionCreationInformation);
        }
    }
}
