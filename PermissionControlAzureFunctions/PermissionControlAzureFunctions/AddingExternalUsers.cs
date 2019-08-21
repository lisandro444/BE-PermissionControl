using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Security;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Newtonsoft.Json;
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
                SiteProperties siteProps = tenant.GetSitePropertiesByUrl(siteUrl, true);
                Site site = tenant.GetSiteByUrl(siteUrl);
                Web web = site.RootWeb;
                Group group = null;
                RoleDefinitionCollection permissionLevels = web.RoleDefinitions;

                ctx.Load(web);
                ctx.Load(web.SiteGroups);
                ctx.Load(siteProps);
                ctx.Load(permissionLevels);

                ctx.ExecuteQuery();  

                if (CheckUserDomainFrom(siteProps, currentEmail))
                {
                    // If group doesn't exist in web, add it
                    if (!GroupExistsInWebSite(web, groupName))
                    {
                        if (groupName == "SCJ External Contribute")
                        {
                            //var permissionLevelExist = permissionLevels.Select(p => p.Name == "SCJ External Contribute").Count();
                            // Create Custom Permission Level
                            //if (permissionLevelExist!=0)
                            CreateContributePermissionLevel(web);
                            // Create new Group
                            group = AddGroup(web, groupName);
                            // Add Custom Pemission Level to Group
                            web.AddPermissionLevelToGroup(groupName, "SCJ External Contribute", true);

                            //web.AddUserToGroup(group, currentEmail);
                        }

                        if (groupName == "SCJ External Read")
                        {
                            // Create Custom Permission Level
                            CreateReadPermissionLevel(web);
                            // Create new Group
                            group = AddGroup(web, groupName);
                            // Add Custom Pemission Level to Group
                            web.AddPermissionLevelToGroup(groupName, "SCJ External Read", true);

                            //web.AddUserToGroup(group, currentEmail);
                        }

                    }
                    else // Just Add the user to group
                    {
                        group = web.SiteGroups.GetByName(groupName);
                        //web.AddUserToGroup(group, "lisandrorossi444@gmail.com");
                        //group.InviteExternalUser("lisandrorossi444@gmail.com");  
                    }
                    ctx.ExecuteQuery();
                    return req.CreateResponse(HttpStatusCode.OK, true);
                }

                return req.CreateResponse(HttpStatusCode.OK, false);

            }
            catch (Exception e)
            {
                return req.CreateResponse(HttpStatusCode.InternalServerError, e.Message);
            }
        }

        private static bool CheckUserDomainFrom(SiteProperties siteProps, string emailAddress)
        {
            var allowedListDomainFromSite = siteProps.SharingAllowedDomainList.Split(',').Select(x => x.Trim().ToUpper()).ToList();
            MailAddress address = new MailAddress(emailAddress);
            string host = address.Host.ToUpper();

            if (allowedListDomainFromSite.Contains(host))
            {
                return true;
            }
            return false;
        }

        private static bool GroupExistsInWebSite(Web web, string name)
        {
            return web.SiteGroups.OfType<Group>().Count(g => g.Title.Equals(name, StringComparison.InvariantCultureIgnoreCase)) > 0;
        }

        private static Group AddGroup(Web web, string groupName)
        {
            var newGroup = web.AddGroup(groupName, "Permission Control - Custom Contribute Group for External User", true, true, false);
            return newGroup;
        }

        private static void CreateContributePermissionLevel(Web web)
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
            perms.Set(PermissionKind.BrowseDirectories);
            perms.Set(PermissionKind.BrowseUserInfo);
            perms.Set(PermissionKind.UseClientIntegration);
            perms.Set(PermissionKind.Open);
            perms.Set(PermissionKind.EditMyUserInfo);
            perms.Set(PermissionKind.ManagePersonalViews);
            perms.Set(PermissionKind.AddDelPrivateWebParts);
            perms.Set(PermissionKind.UpdatePersonalWebParts);
            roleDefinitionCreationInformation.BasePermissions = perms;
            roleDefinitionCreationInformation.Name = "SCJ External Contribute";
            roleDefinitionCreationInformation.Description = "Custom Permission Level - SCJ External Contribute";
            web.RoleDefinitions.Add(roleDefinitionCreationInformation);

        }

        private static void CreateReadPermissionLevel(Web web)
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
            perms.Set(PermissionKind.ViewFormPages);
            perms.Set(PermissionKind.ViewPages);
            perms.Set(PermissionKind.BrowseUserInfo);
            perms.Set(PermissionKind.UseClientIntegration);
            perms.Set(PermissionKind.Open);
            roleDefinitionCreationInformation.BasePermissions = perms;
            roleDefinitionCreationInformation.Name = "SCJ External Read";
            roleDefinitionCreationInformation.Description = "Custom Permission Level - SCJ External Read";
            web.RoleDefinitions.Add(roleDefinitionCreationInformation);
        }

        private static void AddExternalUserAPI(Web web, string siteUrl, string externalUserEmail)
        {
            var emailBody = "Welcome to site " + web.Title;

            var peoplePickerInput = new
            {
                Key = externalUserEmail,
                DisplayText = externalUserEmail,
                IsResolved = true,
                Description = externalUserEmail,
                EntityType = "",
                EntityData = new
                {
                    SPUserID = externalUserEmail,
                    Email = externalUserEmail,
                    IsBlocked = "False",
                    PrincipalType = "UNVALIDATED_EMAIL_ADDRESS",
                    AccountName = externalUserEmail,
                    SIPAddress = externalUserEmail,
                    IsBlockedOnODB = "False"
                },
                MultipleMatches = "",
                ProviderName = "",
                ProviderDisplayName = ""
            };

            var peoplePickerInputSerialized = JsonConvert.SerializeObject(peoplePickerInput);


            var body = new
            {
                emailBody,
                includeAnonymousLinkInEmail = "false",
                peoplePickerInputSerialized,
                roleValue = "group:4",
                sendEmail = "false",
                url = siteUrl,
                useSimplifiedRoles = "true"
            };

            var bodySerialized = JsonConvert.SerializeObject(body);

            var buffer = System.Text.Encoding.UTF8.GetBytes(bodySerialized);
            var byteContent = new ByteArrayContent(buffer);

            byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            var client = new HttpClient();
            var result = client.PostAsync(siteUrl + "/_api/ Web.ShareObject", byteContent).Result;
        }
    }
}
