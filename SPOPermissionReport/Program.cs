/**
Disclaimer:
This Sample Code, scripts or any related information are provided for the purpose of illustration only and is not intended to be used in a production environment.
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY
AND/OR FITNESS FOR A PARTICULAR PURPOSE.We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code,
provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software
product in which the Sample Code is embedded; and(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees,
that arise or result from the use or distribution of the Sample Code.#>
The lambda expressions are inspired from this article: http://www.morgantechspace.com/2017/09/get-item-level-permissions-sharepoint-csom.html

Author:Srinivas Varukala, svarukal@microsoft.com
Date: 2/20/2018
**/

using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Net;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Search.Query;

namespace SPOPermissionReport
{
    class Program
    {
        public static string filePath; // = @"C:\temp\notes.csv";
        public static string removeUsersLogFilePath;
        public static string[] filterUsers;
        public static string[] ignoreLists;
        public static string adminUser;
        public static bool removeFilteredUsers = false;
        public static List<int> siteGroupIDs;
        public static bool skipLists = false;
        public static bool skipListItems = false;
        public static string odbUrlsCsvFile;
        static void Main(string[] args)
        {
            #region READ_SETTINGS
            string adminSiteUrl = ConfigurationManager.AppSettings["adminsite"];
            adminUser = ConfigurationManager.AppSettings["adminuser"];
            filePath = ConfigurationManager.AppSettings["csvfilepath"];
            odbUrlsCsvFile = ConfigurationManager.AppSettings["odbUrlsCsvFile"];
            removeUsersLogFilePath = ConfigurationManager.AppSettings["csvremoveusersfiletolog"];
            filterUsers = ConfigurationManager.AppSettings["filterusers"].Split(',');
            ignoreLists = ConfigurationManager.AppSettings["excludelists"].Split(',');
            removeFilteredUsers = Convert.ToBoolean(ConfigurationManager.AppSettings["removefilteredusers"]);
            bool useCSVAsInput = Convert.ToBoolean(ConfigurationManager.AppSettings["UseCSVAsInput"]);
            skipLists = Convert.ToBoolean(ConfigurationManager.AppSettings["skipLists"]);
            skipListItems = Convert.ToBoolean(ConfigurationManager.AppSettings["skipListItems"]);
            string fileWithSiteUrlsToProcess = ConfigurationManager.AppSettings["csvfilesitestoprocess"];
            string sitesScope = ConfigurationManager.AppSettings["sitesScope"];
            string mySiteHost = ConfigurationManager.AppSettings["mySiteHost"]; 
            string adminMySiteUrl = ConfigurationManager.AppSettings["AdminODBUrl"];

            var pwdtxt = "**********";
            Console.WriteLine("---------------------------------------");
            Console.WriteLine("Environment details provided\n");
            Console.WriteLine("Admin url:                           {0}", adminSiteUrl);
            Console.WriteLine("Admin user:                          {0}", adminUser);
            Console.WriteLine("Site Urls Input:                     {0}", useCSVAsInput ? fileWithSiteUrlsToProcess : "All SPO and ODB Sites");
            if (!useCSVAsInput)
            {
            Console.WriteLine("Limit sites to be processed to:      {0}", sitesScope + " Sites");
                if(sitesScope.Equals("ODB"))
                    Console.WriteLine("ODB Sites/MySite Host URL:           {0}", mySiteHost);
            }
            Console.WriteLine("Skip lists (true or false)?:         {0}", skipLists ? "true" : "false");
            Console.WriteLine("Skip list items (true or false)?:    {0}", skipListItems ? "true" : "false");

            Console.WriteLine("Users/Groups to be idenified:        {0}", string.Join("|", filterUsers.ToArray()));
            Console.WriteLine("Lists to be excluded:                {0}", string.Join("|", ignoreLists.ToArray()));
            Console.WriteLine("CSV Log File to log filtered perms:  {0}", filePath);
            
            Console.WriteLine("Remove users (true or false)?:       {0}", removeFilteredUsers? "true" : "false");
            if(removeFilteredUsers)
            {
            Console.WriteLine("CSV Log File to log removed users:   {0}", removeUsersLogFilePath);
            }
            Console.WriteLine("---------------------------------------");

            Console.Write("Are you sure you want to proceed? [y/n] ");
            ConsoleKey k = Console.ReadKey(false).Key;
            if(k != ConsoleKey.Y)
            {
                Console.WriteLine("\nHit enter to cancel.");
                Console.Read();
                return;
            }
            #endregion READ_SETTINGS

            #region Get password securely

            pwdtxt = string.Empty;
            Console.WriteLine("\nEnter your password for the admin account: ");
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pwdtxt += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && pwdtxt.Length > 0)
                    {
                        pwdtxt = pwdtxt.Substring(0, (pwdtxt.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            }
            // Stops Receving Keys Once Enter is Pressed
            while (key.Key != ConsoleKey.Enter);

            Console.WriteLine();
            //Console.WriteLine("The Password You entered is : " + pwdtxt);
            

            //pwdtxt = "4851@svarukal";
            if(pwdtxt.Trim().Length == 0)
            {
                Console.WriteLine("Password is empty. Please try again. Press enter to close.");
                return;
            }
            SecureString pwd = new SecureString();
            foreach (char c in pwdtxt.ToCharArray())
                pwd.AppendChar(c);
            #endregion

            #region COMMENTED_TestODBSitesProcess
            Console.Write("Enter 'y' to write all ODB urls to csv file? [y/n] ");
            ConsoleKey g = Console.ReadKey(false).Key;
            if (g == ConsoleKey.Y)
            {
                using (ClientContext adminCtx1 = new ClientContext(adminSiteUrl))
                {
                    SharePointOnlineCredentials creds1 = new SharePointOnlineCredentials(adminUser, pwd);
                    adminCtx1.Credentials = creds1;
                    GetUsersODBSites(adminCtx1);
                    Console.WriteLine("ODB Site Urls written to file: {0}\n Hit enter to close", odbUrlsCsvFile);
                    Console.Read();
                    return;
                }
            }
            else
                Console.WriteLine("\nMoving on...");

            /*
            List<string> users = GetUserAccountNames(adminCtx1);
            foreach (var user in users)
            {
                Console.WriteLine(user);
            }

            UserProfileService uprofService = new UserProfileService();

            uprofService.Url = adminSiteUrl + "/_vti_bin/UserProfileService.asmx";
            uprofService.UseDefaultCredentials = false;

            Uri targetSite = new Uri(adminSiteUrl);
            uprofService.CookieContainer = new CookieContainer();
            string authCookieValue = creds1.GetAuthenticationCookie(targetSite);
            uprofService.CookieContainer.SetCookies(targetSite, authCookieValue);
            var userProfileResult1 = uprofService.GetUserProfileByIndex(-1);
            long numProfiles = uprofService.GetUserProfileCount();
            Console.WriteLine("Total user profiles: " + numProfiles);
            while (userProfileResult1.NextValue != "-1")
            {
                string personalUrl = null;
                foreach (var u in userProfileResult1.UserProfile)
                {
                    if (u.Values.Length != 0 && u.Values[0].Value != null && u.Name == "PersonalSpace")
                    {
                        personalUrl = u.Values[0].Value as string;
                        Console.WriteLine(mySiteHost + personalUrl);
                        //BeginProcessWeb("", creds1);
                        break;
                    }
                }
                int nextIndex = -1;
                nextIndex = Int32.Parse(userProfileResult1.NextValue);
                userProfileResult1 = uprofService.GetUserProfileByIndex(nextIndex);
            }
            Console.ReadKey();
            return;
            */
            #endregion


            //string filePath = @"C:\temp\notes.csv";
            string txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", "Site Title", "Site URL", "List Title", "Path", "ObjectType", "Principal Type", "Principal ID", "Permission Type", "Notes", Environment.NewLine);
            if (removeFilteredUsers)
                System.IO.File.AppendAllText(removeUsersLogFilePath, txt, ASCIIEncoding.ASCII);
            else
                System.IO.File.AppendAllText(filePath, txt, ASCIIEncoding.ASCII);
            SharePointOnlineCredentials creds = new SharePointOnlineCredentials(adminUser, pwd);
            ClientContext adminCtx = new ClientContext(adminSiteUrl);
            adminCtx.Credentials = creds;
            if (useCSVAsInput)
            {
                #region ReadCSVInput
                //SharePointOnlineCredentials creds = new SharePointOnlineCredentials(adminUser, pwd);
                List<string> sitesList = new List<string>();
                try
                {
                    using (var reader = new StreamReader(fileWithSiteUrlsToProcess))
                    {
                        while (!reader.EndOfStream)
                        {
                            sitesList.Add(reader.ReadLine().Trim());
                        }
                    }
                }
                catch(Exception e1)
                {
                    Console.WriteLine("Error reading csv files contains site urls to process. Error details {0}", e1.ToString());
                    Console.Read();
                    return;
                }
                foreach (string siteurl in sitesList)
                {
                    if (siteurl.Contains(mySiteHost))
                    {
                        AddRemoveODBSiteAdmin(siteurl, true, adminCtx);
                        BeginProcessWeb(siteurl, creds);
                        AddRemoveODBSiteAdmin(siteurl, false, adminCtx);
                    }
                    else
                        BeginProcessWeb(siteurl, creds);
                }
                #endregion
            }
            else
            {
                #region READ_USING_API
                //SharePointOnlineCredentials creds = new SharePointOnlineCredentials(adminUser, pwd);
                if (sitesScope.Equals("All") || sitesScope.Equals("SPO"))
                {
                    Tenant tenant = new Tenant(adminCtx);
                    int startIndex = 0;
                    SPOSitePropertiesEnumerable allSites = null;

                    while (allSites == null || allSites.Count > 0)
                    {
                        allSites = tenant.GetSiteProperties(startIndex, false);
                        adminCtx.Load(allSites);
                        adminCtx.ExecuteQuery();
                        if (allSites != null && allSites.Count > 0)
                        {
                            foreach (SiteProperties prop in allSites)
                            {
                                BeginProcessWeb(prop.Url, creds);
                            }
                            startIndex += allSites.Count;
                        }
                    }
                }
                if (sitesScope.Equals("All") || sitesScope.Equals("ODB"))
                {
                    #region Process SPO and ODB
                    //ClientContext adminCtx1 = new ClientContext(adminSiteUrl);
                    //SharePointOnlineCredentials creds1 = new SharePointOnlineCredentials(adminUser, pwd);
                    UserProfileService userProfileSvc = new UserProfileService();
                    //PeopleManager pplMgr = new PeopleManager(adminCtx);

                    userProfileSvc.Url = adminSiteUrl + "/_vti_bin/UserProfileService.asmx";
                    userProfileSvc.UseDefaultCredentials = false;
                    userProfileSvc.CookieContainer = new CookieContainer();
                    string authCookie = creds.GetAuthenticationCookie(new Uri(adminSiteUrl));
                    userProfileSvc.CookieContainer.SetCookies(new Uri(adminSiteUrl), authCookie);
                    var userProfileResult = userProfileSvc.GetUserProfileByIndex(-1);
                    long totalProfiles = userProfileSvc.GetUserProfileCount();
                    Console.WriteLine("\nTotal user profiles: " + totalProfiles);
                    string personalUrl = null;
                    int nextIndex = -1;
                    while (userProfileResult.NextValue != "-1")
                    {
                        foreach (var u in userProfileResult.UserProfile)
                        {
                            /* (PersonalSpace is the name of the path to a user's OneDrive for Business site. Users who have not yet created a OneDrive for Business site might not have this property set.)*/
                            if (u.Values.Length != 0 && u.Values[0].Value != null && u.Name == "PersonalSpace")
                            {
                                try
                                {
                                    personalUrl = u.Values[0].Value as string;
                                    AddRemoveODBSiteAdmin(mySiteHost + personalUrl, true, adminCtx);
                                    BeginProcessWeb(mySiteHost + personalUrl, creds);
                                    AddRemoveODBSiteAdmin(mySiteHost + personalUrl, false, adminCtx);
                                    break;
                                }
                                catch(Exception odbErr)
                                {
                                    Console.WriteLine("Exception while add/remove of Site Admin to DDB site collection {0}. Error details: {1}", u.Values[0].Value as string, odbErr.ToString());
                                }
                            }
                        }
                        nextIndex = Int32.Parse(userProfileResult.NextValue);
                        userProfileResult = userProfileSvc.GetUserProfileByIndex(nextIndex);
                    }
                    #endregion
                }
                
                #endregion
            }
            adminCtx.Dispose();
            Console.WriteLine("Finished processing all site collections.");
            Console.ReadLine();
            Console.Write("Are you sure you want to close it? [y/n] ");
            ConsoleKey r = Console.ReadKey(false).Key;
            if (r != ConsoleKey.Y)
            {
                Console.WriteLine("\nHit enter to close.");
                Console.Read();
                return;
            }
        }

        /// <summary>
        /// This method is not used as of now
        /// </summary>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        private static List<string> GetUserAccountNames(ClientContext clientContext)
        {
            //Use search to get the AccountNames of users in a tenant. 
            var keywordQuery = new KeywordQuery(clientContext);
            keywordQuery.QueryText = "*";
            keywordQuery.SourceId = new Guid("b09a7990-05ea-4af9-81ef-edfab16c4e31");
            keywordQuery.SelectProperties.Add("AccountName");
            keywordQuery.BypassResultTypes = true;
            var searchExecutor = new SearchExecutor(clientContext);
            var searchResults = searchExecutor.ExecuteQuery(keywordQuery);
            clientContext.ExecuteQuery();
            var accountNames = new List<string>();
            if (searchResults.Value.Count > 0 && searchResults.Value[0].RowCount > 0)
            {
                foreach (var resultRow in searchResults.Value[0].ResultRows)
                {
                    accountNames.Add(resultRow["AccountName"].ToString());
                }

            }
            return accountNames;
        }
        private static void GetUsersODBSites(ClientContext clientContext)
        {
            Console.WriteLine("\nGetting user profiles...\n");
            //Use search to get the AccountNames of users in a tenant. 
            var keywordQuery = new KeywordQuery(clientContext);
            keywordQuery.QueryText = "*";
            keywordQuery.SourceId = new Guid("b09a7990-05ea-4af9-81ef-edfab16c4e31");
            keywordQuery.SelectProperties.Add("AccountName");
            keywordQuery.BypassResultTypes = true;
            var searchExecutor = new SearchExecutor(clientContext);
            var searchResults = searchExecutor.ExecuteQuery(keywordQuery);
            clientContext.ExecuteQuery();
            var accountNames = new List<string>();
            Console.WriteLine("Found {0} user profiles.\n", searchResults.Value[0].RowCount);
            if (searchResults.Value.Count > 0 && searchResults.Value[0].RowCount > 0)
            {
                PeopleManager peopleManager = new PeopleManager(clientContext);
                var odbSiteUrl = ConfigurationManager.AppSettings["mySiteHost"];
                foreach (var resultRow in searchResults.Value[0].ResultRows)
                {
                    accountNames.Add(resultRow["AccountName"].ToString());
                    var odbObj = peopleManager.GetUserProfilePropertyFor(resultRow["AccountName"].ToString(), "PersonalSpace");
                    clientContext.ExecuteQuery();
                    if (!string.IsNullOrEmpty(odbObj.Value))
                    {
                        Console.WriteLine(odbSiteUrl + odbObj.Value);
                        System.IO.File.AppendAllText(odbUrlsCsvFile, odbSiteUrl + odbObj.Value + Environment.NewLine, ASCIIEncoding.ASCII);
                    }
                    else
                        Console.WriteLine("..");
                }
            }

            //return accountNames;
        }


        public static string adminLoginName = null;
        public static void AddRemoveODBSiteAdmin(string odbUrl, bool add, ClientContext adminCtx)
        {
            if (odbUrl.IndexOf(ConfigurationManager.AppSettings["AdminODBUrl"].ToString(), StringComparison.CurrentCultureIgnoreCase) >= 0)
                return;
            var tenant = new Tenant(adminCtx);
            if (string.IsNullOrEmpty(adminLoginName))
            { 
                var currentUser = adminCtx.Web.CurrentUser;
                adminCtx.Load(currentUser, u => u.LoginName);
                adminCtx.ExecuteQuery();
                adminLoginName = currentUser.LoginName;
            }
            tenant.SetSiteAdmin(odbUrl, adminLoginName, add);
            adminCtx.ExecuteQuery();
            #region Commented
            /*
            using (var context = new ClientContext(odbUrl))
            {
                //creds.UserName
                context.Credentials = creds;
                var rootWeb = context.Web;
                context.Load(rootWeb);
                var spUser = rootWeb.EnsureUser(adminUser);
                context.Load(spUser);
                context.ExecuteQuery();

                if (add && !spUser.IsSiteAdmin)
                {
                    spUser.IsSiteAdmin = true;
                    spUser.Update();
                    context.Load(spUser);
                    context.ExecuteQuery();
                }
                else if(!add && spUser.IsSiteAdmin)
                {
                    spUser.IsSiteAdmin = false;
                    spUser.Update();
                    context.Load(spUser);
                    context.ExecuteQuery();

                }
            }
            */
            #endregion
        }
        public static void BeginProcessWeb(string siteUrl, SharePointOnlineCredentials creds)
        {
            try
            {
                Console.WriteLine("\nSite Collection:\t" + siteUrl);
                using (var context = new ClientContext(siteUrl))
                {
                    context.Credentials = creds;
                    var rootWeb = context.Web;
                    var subWebs = rootWeb.Webs;
                    context.Load(rootWeb);
                    context.Load(subWebs, w => w.Include(a => a.HasUniqueRoleAssignments, a => a.Url));
                    context.ExecuteQuery();
                    ProcessWeb(rootWeb, creds, true);
                    GetAllWebs(subWebs, creds);
                }
            }
            catch (Exception e1)
            {
                Console.WriteLine("Exception while processing site collection {0}. Error details: {1}", siteUrl, e1.ToString());
                System.IO.File.AppendAllText(filePath, string.Format("\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", "n/a", SanitizeForCSV(siteUrl), "n/a", "n/a", "Web", "n/a", "n/a", "n/a", "ERROR( " + SanitizeForCSV(e1.Message) + " )", Environment.NewLine), ASCIIEncoding.ASCII);
            }
        }
        public static void GetAllWebs(WebCollection webs, SharePointOnlineCredentials creds)
        {
            foreach (var web in webs)
            {
                try
                { 
                    ProcessWeb(web, creds, false);
                    using (var context = new ClientContext(web.Url))
                    {
                        context.Credentials = creds;
                        var subwebs = context.Web.Webs;
                        context.Load(subwebs, w => w.Include(a => a.HasUniqueRoleAssignments, a=>a.Url));
                        context.ExecuteQuery();
                        if (subwebs.Count > 0)
                            GetAllWebs(subwebs, creds);
                    }
                }
                catch (Exception e1)
                {
                    Console.WriteLine("Exception while processing web {0}. Error details: {1}", web.Url, e1.ToString());
                    //System.IO.File.AppendAllText(filePath, string.Format("\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", "n/a", SanitizeForCSV(web.ur), "n/a", "n/a", "Web", "n/a", "n/a", "n/a", "ERROR( " + SanitizeForCSV(e1.Message) + " )", Environment.NewLine), ASCIIEncoding.ASCII);
                }
            }
        }
        public static void ProcessWeb(Web subweb, SharePointOnlineCredentials creds, bool isRootWeb)
        {
            using (var context = new ClientContext(subweb.Url))
            {
                context.Credentials = creds;
                var web = context.Web;
                context.Load(web, w => w.Title, w => w.Url,w => w.HasUniqueRoleAssignments, w=>w.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
                                  roleAsg => roleAsg.Member.Id,
                                  roleAsg => roleAsg.Member.PrincipalType,
                                  roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name,
                                  roleDef => roleDef.Description)));
                context.ExecuteQuery();
                Console.WriteLine("\tWeb: {0},  IsRootWeb: {1}", web.Title, isRootWeb);

                if (isRootWeb)
                {
                    if (siteGroupIDs != null)
                        siteGroupIDs.Clear();
                    else
                        siteGroupIDs = new List<int>();
                }
                if(web.HasUniqueRoleAssignments)
                    ProcessUniquePerms(context, web, isRootWeb);

                #region PROCESS LISTS AND LIST ITEMS
                if (!skipLists)
                {
                    context.Load(web.Lists, a => a.IncludeWithDefaultProperties(b => b.HasUniqueRoleAssignments),
                        permsn => permsn.Include(a => a.Title, a => a.Id, a => a.RootFolder, a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
                                      roleAsg => roleAsg.Member.Id,
                                      roleAsg => roleAsg.Member.PrincipalType,
                                      roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name,
                                      roleDef => roleDef.Description))));
                    context.ExecuteQuery();
                    foreach (List list in web.Lists)
                    {
                        try
                        {
                            if (!ignoreLists.Contains(list.Title, StringComparer.CurrentCultureIgnoreCase) && !list.Hidden)
                            {
                                Console.WriteLine("\t\tList title:{0}", list.Title);
                                //List list = web.Lists.GetByTitle("Documents");

                                if (list.HasUniqueRoleAssignments)
                                {
                                    Console.WriteLine("\t\t\t{0} has unique perms", list.Title);
                                    ProcessUniquePerms(context, list);
                                }

                                #region PROCESS LIST ITEMS
                                if (!skipListItems)
                                {
                                    string query = @"
                                                    <View Scope='RecursiveAll'>
                                                        <Query>
                                                            <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
                                                        </Query>
                                                        <RowLimit Paged='TRUE'>5000</RowLimit>
                                                    </View>
                                                    ";
                                    ListItemCollectionPosition position = null;
                                    List<ListItem> allItems = new List<ListItem>();
                                    ListItemCollection listItems = null;
                                    int i = 0;
                                    do
                                    {
                                        var camlQuery = new CamlQuery();
                                        camlQuery.ListItemCollectionPosition = position;
                                        camlQuery.ViewXml = query;
                                        listItems = list.GetItems(camlQuery);
                                        context.Load(listItems, l => l.ListItemCollectionPosition, a => a.IncludeWithDefaultProperties(b => b.HasUniqueRoleAssignments, b => b.FileSystemObjectType),
                                            permsn => permsn.Include(a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
                                                    roleAsg => roleAsg.Member.Id,
                                                    roleAsg => roleAsg.Member.PrincipalType,
                                                    roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name,
                                                    roleDef => roleDef.Description))));
                                        context.ExecuteQuery();
                                        position = listItems.ListItemCollectionPosition;
                                        allItems.AddRange(listItems.Where(li => li.HasUniqueRoleAssignments == true));
                                        i++;
                                    }
                                    while (position != null);
                                    //if (i > 0)
                                    //  Console.WriteLine("Iteration count:" + i);
                                    Console.WriteLine("\t\t\tFound {0} unique permissioned list items,files,folders", allItems.Count);
                                    if (allItems != null)
                                    {
                                        foreach (var item in allItems)
                                        {
                                            ProcessUniquePerms(context, item, list);
                                        }
                                    }
                                }
                                #endregion
                            }
                        }
                        catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException noAccessEx)
                        {
                            Console.WriteLine("No access to list: " + list.Title);
                        }
                        catch (Exception ex1)
                        {
                            Console.WriteLine("Unknown error processing list: " + list.Title + ". Error details: " + ex1.ToString());
                        }
                        //Console.WriteLine("-------------------------------------");
                    }
                }
                #endregion 
            }
        }
        public static string ProcessLoginName(string memberLoginName)
        {
            if(memberLoginName.Equals("c:0(.s|true"))
            {
                return "Everyone";
            }
            else if(memberLoginName.Contains("spo-grid-all-users"))
            {
                return "Everyone except external users";
            }
            return memberLoginName;
        }
        public static string SanitizeForCSV(string data)
        {
            return data.Replace(',', '_');
        }
        public static void ProcessUniquePerms(ClientContext context, ListItem item, List list)
        {
            string txt = string.Empty;
            List<RoleAssignment> roleAsgs = new List<RoleAssignment>();
            foreach (var roleAsg in item.RoleAssignments)
            {
                //Console.WriteLine("User/Group: " + roleAsg.Member.LoginName);
                List<string> roles = new List<string>();
                foreach (var role in roleAsg.RoleDefinitionBindings)
                {
                    roles.Add(role.Name);
                }
                //Console.WriteLine("  Permissions: " + string.Join(",", roles.ToArray()));
                string perms = "\"" + string.Join("|", roles.ToArray()) + "\"";
                string pId = ProcessLoginName(roleAsg.Member.LoginName);
                if (filterUsers.Contains(pId, StringComparer.CurrentCultureIgnoreCase))
                {
                    if (removeFilteredUsers)
                    {
                        roleAsgs.Add(roleAsg);
                    }
                    else
                    {
                        txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, pId, string.Join("|", roles.ToArray()), "Granted directly", Environment.NewLine);
                        System.IO.File.AppendAllText(filePath, txt, ASCIIEncoding.ASCII);
                    }
                }
                if (roleAsg.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup)
                {
                    //Process only if the web has unique SPGroup i.e do not process groups inheirted from Rootweb
                    if (!siteGroupIDs.Contains(roleAsg.Member.Id))
                    {
                        List<User> usersToRemove = new List<User>();
                        var grp = context.Web.SiteGroups.GetByName(roleAsg.Member.LoginName);
                        context.Load(grp, g => g.Users);
                        context.ExecuteQuery();
                        //Console.WriteLine("  User count: " + grp.Users.Count);
                        foreach (var user in grp.Users)
                        {
                            pId = string.Empty;
                            pId = ProcessLoginName(user.LoginName);
                            if (filterUsers.Contains(pId, StringComparer.CurrentCultureIgnoreCase))
                            {
                                if (removeFilteredUsers)
                                {
                                    usersToRemove.Add(user);
                                }
                                else
                                {
                                    txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, pId, string.Join("|", roles.ToArray()), "Granted  Through Group Membership. Group name:" + roleAsg.Member.LoginName, Environment.NewLine);
                                    System.IO.File.AppendAllText(filePath, txt, ASCIIEncoding.ASCII);
                                }
                            }
                        }
                        if (usersToRemove.Count > 0)
                        {
                            foreach (var user in usersToRemove)
                            {
                                try
                                {
                                    grp.Users.Remove(user);
                                    grp.Update();
                                    context.ExecuteQuery();
                                    txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, user.LoginName, "", "User removed from group: " + roleAsg.Member.LoginName, Environment.NewLine);
                                    System.IO.File.AppendAllText(removeUsersLogFilePath, txt, ASCIIEncoding.ASCII);
                                }
                                catch (Exception e2)
                                {
                                    Console.WriteLine("Failed to remove user {0} from group {1}. Error details: {2}", user.LoginName, grp.LoginName, e2.ToString());
                                }
                            }
                        }
                    }
                }
                //Console.WriteLine(Environment.NewLine);
            }
            if (roleAsgs.Count > 0)
            {
                try
                {
                    foreach (var roleAsg in roleAsgs)
                    {
                        roleAsg.RoleDefinitionBindings.RemoveAll();
                        roleAsg.DeleteObject();
                        context.ExecuteQuery();
                        txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, ProcessLoginName(roleAsg.Member.LoginName), "", "User/Group removed", Environment.NewLine);
                        System.IO.File.AppendAllText(removeUsersLogFilePath, txt, ASCIIEncoding.ASCII);
                    }
                }
                catch (Exception e1)
                {
                    Console.WriteLine("Failed to remove user/group from list item: {0}. Error details: {1}", item["FileRef"], e1.ToString());
                }
            }
        }
        public static void ProcessUniquePerms(ClientContext context, List list)
        {
            string txt = string.Empty;
            List<RoleAssignment> roleAsgs = new List<RoleAssignment>();
            foreach (var roleAsg in list.RoleAssignments)
            {
                //Console.WriteLine("User/Group: " + roleAsg.Member.LoginName);
                List<string> roles = new List<string>();
                foreach (var role in roleAsg.RoleDefinitionBindings)
                {
                    roles.Add(role.Name);
                }
                //Console.WriteLine("  Permissions: " + string.Join(",", roles.ToArray()));
                string perms = "\"" + string.Join("|", roles.ToArray()) + "\"";
                string pId = ProcessLoginName(roleAsg.Member.LoginName);
                if (filterUsers.Contains(pId, StringComparer.CurrentCultureIgnoreCase))
                {
                    if (removeFilteredUsers)
                    {
                        roleAsgs.Add(roleAsg);
                    }
                    else
                    {
                        txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, pId, string.Join("|", roles.ToArray()), "Granted directly", Environment.NewLine);
                        System.IO.File.AppendAllText(filePath, txt, ASCIIEncoding.ASCII);
                    }
                }
                if (roleAsg.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup)
                {
                    //Process only if the list has unique SPGroup i.e do not process groups inheirted from Rootweb
                    if (!siteGroupIDs.Contains(roleAsg.Member.Id))
                    {
                        List<User> usersToRemove = new List<User>();
                        var grp = context.Web.SiteGroups.GetByName(roleAsg.Member.LoginName);
                        context.Load(grp, g => g.Users);
                        context.ExecuteQuery();
                        //Console.WriteLine("  User count: " + grp.Users.Count);
                        foreach (var user in grp.Users)
                        {
                            pId = string.Empty;
                            pId = ProcessLoginName(user.LoginName);
                            if (filterUsers.Contains(pId, StringComparer.CurrentCultureIgnoreCase))
                            {
                                if (removeFilteredUsers)
                                {
                                    usersToRemove.Add(user);
                                }
                                else
                                {
                                    txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, pId, string.Join("|", roles.ToArray()), "Granted  Through Group Membership. Group name:" + roleAsg.Member.LoginName, Environment.NewLine);
                                    System.IO.File.AppendAllText(filePath, txt, ASCIIEncoding.ASCII);
                                }
                            }
                        }
                        if (usersToRemove.Count > 0)
                        {
                            foreach (var user in usersToRemove)
                            {
                                try
                                {
                                    grp.Users.Remove(user);
                                    grp.Update();
                                    context.ExecuteQuery();
                                    txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, user.LoginName, "", "User removed from group: " + roleAsg.Member.LoginName, Environment.NewLine);
                                    System.IO.File.AppendAllText(removeUsersLogFilePath, txt, ASCIIEncoding.ASCII);
                                }
                                catch (Exception e2)
                                {
                                    Console.WriteLine("Failed to remove user {0} from group {1}. Error details: {2}", user.LoginName, grp.LoginName, e2.ToString());
                                }
                            }
                        }
                    }
                }
                //Console.WriteLine(Environment.NewLine);
            }
            if (roleAsgs.Count > 0)
            {
                try
                {
                    foreach (var roleAsg in roleAsgs)
                    {
                        roleAsg.RoleDefinitionBindings.RemoveAll();
                        roleAsg.DeleteObject();
                        context.ExecuteQuery();
                        txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(list.Title), list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, ProcessLoginName(roleAsg.Member.LoginName), "", "User/Group removed", Environment.NewLine);
                        System.IO.File.AppendAllText(removeUsersLogFilePath, txt, ASCIIEncoding.ASCII);
                    }
                }
                catch (Exception e1)
                {
                    Console.WriteLine("Failed to remove user/group from list: {0}. Error details: {1}", list.Title, e1.ToString());
                }
            }
        }
        public static void ProcessUniquePerms(ClientContext context, Web web, bool isRootWeb)
        {
            string txt = string.Empty;
            List<RoleAssignment> roleAsgs = new List<RoleAssignment>();
            //foreach (var roleAsg in web.RoleAssignments) ;
            int totalRolesAssignments = web.RoleAssignments.Count;
            for (var i = 0; i < totalRolesAssignments; i++)
            {
                var roleAsg = web.RoleAssignments[i];
                //Console.WriteLine("User/Group: " + roleAsg.Member.LoginName);
                List<string> roles = new List<string>();
                foreach (var role in roleAsg.RoleDefinitionBindings)
                {
                    roles.Add(role.Name);
                }
                //Console.WriteLine("  Permissions: " + string.Join(",", roles.ToArray()));
                string perms = "\"" + string.Join("|", roles.ToArray()) + "\"";
                string pId = ProcessLoginName(roleAsg.Member.LoginName);
                if (filterUsers.Contains(pId, StringComparer.CurrentCultureIgnoreCase))
                {
                    if (removeFilteredUsers)
                    {
                        roleAsgs.Add(roleAsg);
                    }
                    else
                    {
                        txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(web.Title), "n/a", "Web", roleAsg.Member.PrincipalType, pId, string.Join("|", roles.ToArray()), "Granted directly", Environment.NewLine);
                        System.IO.File.AppendAllText(filePath, txt, ASCIIEncoding.ASCII);
                    }
                }
                if (roleAsg.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup)
                {
                    //Process only if the web has unique SPGroup i.e do not process groups inheirted from Rootweb
                    if (!siteGroupIDs.Contains(roleAsg.Member.Id))
                    {
                        List<User> usersToRemove = new List<User>();
                        var grp = context.Web.SiteGroups.GetByName(roleAsg.Member.LoginName);
                        context.Load(grp, g => g.Users, g => g.Id);
                        context.ExecuteQuery();
                        if (isRootWeb) siteGroupIDs.Add(grp.Id);
                        //Console.WriteLine("Groupname: {0}, ID: {1}, RoleAsg_ID: {2}, User count: {3}", roleAsg.Member.LoginName, grp.Id, roleAsg.Member.Id, grp.Users.Count);
                        foreach (var user in grp.Users)
                        {
                            pId = string.Empty;
                            pId = ProcessLoginName(user.LoginName);
                            if (filterUsers.Contains(pId, StringComparer.CurrentCultureIgnoreCase))
                            {
                                if (removeFilteredUsers)
                                {
                                    usersToRemove.Add(user);
                                }
                                else
                                {
                                    txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, SanitizeForCSV(web.Title), "n/a", "Web", roleAsg.Member.PrincipalType, pId, string.Join("|", roles.ToArray()), "Granted  Through Group Membership. Group name:" + roleAsg.Member.LoginName, Environment.NewLine);
                                    System.IO.File.AppendAllText(filePath, txt, ASCIIEncoding.ASCII);
                                }


                            }
                        }
                        if (usersToRemove.Count > 0)
                        {
                            foreach (var user in usersToRemove)
                            {
                                try
                                {
                                    grp.Users.Remove(user);
                                    grp.Update();
                                    context.ExecuteQuery();
                                    txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, "Removed user/group", "n/a", "Web", roleAsg.Member.PrincipalType, user.LoginName, "", "User removed from group: " + roleAsg.Member.LoginName, Environment.NewLine);
                                    System.IO.File.AppendAllText(removeUsersLogFilePath, txt, ASCIIEncoding.ASCII);
                                }
                                catch (Exception e2)
                                {
                                    Console.WriteLine("Failed to remove user {0} from group {1}. Error details: {2}", user.LoginName, grp.LoginName, e2.ToString());
                                }
                            }
                        }
                    }
                }
                //Console.WriteLine(Environment.NewLine);
            }

            if(roleAsgs.Count > 0)
            {
                try
                {
                    foreach (var roleAsg in roleAsgs)
                    {
                        roleAsg.RoleDefinitionBindings.RemoveAll();
                        roleAsg.DeleteObject();
                        context.ExecuteQuery();
                        txt = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}{9}", SanitizeForCSV(context.Web.Title), context.Web.Url, "Removed user/group", "n/a", "Web", roleAsg.Member.PrincipalType, ProcessLoginName(roleAsg.Member.LoginName), "", "User/Group removed", Environment.NewLine);
                        System.IO.File.AppendAllText(removeUsersLogFilePath, txt, ASCIIEncoding.ASCII);
                    }
                }
                catch (Exception e1)
                {
                    Console.WriteLine("Failed to remove user/group from site: {0}. Error details: {1}", web.Url, e1.ToString());
                }
            }
        }
    }
}
