/**
Disclaimer:
This Sample Code, scripts or any related information are provided for the purpose of illustration only and is not intended to be used in a production environment.
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY
AND/OR FITNESS FOR A PARTICULAR PURPOSE.We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code,
provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software
product in which the Sample Code is embedded; and(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees,
that arise or result from the use or distribution of the Sample Code.#>
Some of the code is picked from this article: http://www.morgantechspace.com/2017/09/get-item-level-permissions-sharepoint-csom.html

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
		public static bool WhatIf = false;
		static void Main( string[] args )
		{
			string adminSiteUrl = ConfigurationManager.AppSettings["adminsite"];
			adminUser = ConfigurationManager.AppSettings["adminuser"];
			filePath = ConfigurationManager.AppSettings["csvfilepath"];
			removeUsersLogFilePath = ConfigurationManager.AppSettings["csvremoveusersfiletolog"];
			filterUsers = ConfigurationManager.AppSettings["filterusers"].Split( ',' );
			ignoreLists = ConfigurationManager.AppSettings["excludelists"].Split( ',' );
			removeFilteredUsers = Convert.ToBoolean( ConfigurationManager.AppSettings["removefilteredusers"] );
			skipLists = Convert.ToBoolean( ConfigurationManager.AppSettings["skiplists"] );
			WhatIf = Convert.ToBoolean( ConfigurationManager.AppSettings["WhatIf"] );
			bool validateSitesToProcess = Convert.ToBoolean( ConfigurationManager.AppSettings["validatesitestoprocess"] );
			string fileWithSiteUrlsToProcess = ConfigurationManager.AppSettings["csvfilesitestoprocess"];

			//string filePath = @"C:\temp\notes.csv";
			string txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", "Site Title", "Site URL", "List Title", "Path", "ObjectType", "Principal Type", "Principal ID", "Permission Type", "Notes", Environment.NewLine );
			if ( removeFilteredUsers )
			{
				if ( System.IO.File.Exists( removeUsersLogFilePath ) )
					System.IO.File.Delete( removeUsersLogFilePath );
				System.IO.File.AppendAllText( removeUsersLogFilePath, txt, ASCIIEncoding.UTF8 );
			} else
			{
				if ( System.IO.File.Exists( filePath ) )
					System.IO.File.Delete( filePath );
				System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
			}

			var pwdtxt = "**********";

			Console.WriteLine( "---------------------------------------" );
			Console.WriteLine( "Environment details provided\n" );
			Console.WriteLine( "Admin url:                          {0}", adminSiteUrl );
			Console.WriteLine( "Admin user:                         {0}", adminUser );
			Console.WriteLine( "Users to be idenified:              {0}", string.Join( "|", filterUsers.ToArray() ) );
			Console.WriteLine( "Lists to be excluded:               {0}", string.Join( "|", ignoreLists.ToArray() ) );
			Console.WriteLine( "CSV Log File to log filtered perms: {0}", filePath );
			Console.WriteLine( "CSV Log File to log removed users:  {0}", removeUsersLogFilePath );
			Console.WriteLine( "Remove users (true or false)?:      {0}", removeFilteredUsers ? "true" : "false" );
			Console.WriteLine( "Skip Lists (true or false)?:        {0}", skipLists ? "true" : "false" );
			Console.WriteLine( "WhatIf (true or false)?:        {0}", WhatIf ? "true" : "false" );
			Console.WriteLine( "---------------------------------------\n" );
			Console.Write( "Are you sure you want to proceed? [y/n] " );

			ConsoleKey k = Console.ReadKey( false ).Key;
			if ( k != ConsoleKey.Y )
			{
				Console.WriteLine( "\nHit enter to cancel." );
				Console.Read();
				return;
			}
			#region Get password securely
			pwdtxt = string.Empty;
			Console.WriteLine( "\nEnter your password for the admin account: " );
			ConsoleKeyInfo key;

			do
			{
				key = Console.ReadKey( true );

				if ( key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter )
				{
					pwdtxt += key.KeyChar;
					Console.Write( "*" );
				} else
				{
					if ( key.Key == ConsoleKey.Backspace && pwdtxt.Length > 0 )
					{
						pwdtxt = pwdtxt.Substring( 0, (pwdtxt.Length - 1) );
						Console.Write( "\b \b" );
					}
				}
			}
			// Stops Receving Keys Once Enter is Pressed
			while ( key.Key != ConsoleKey.Enter );

			Console.WriteLine();
			//Console.WriteLine("The Password You entered is : " + pwdtxt);
			#endregion

			#region Validate Password
			if ( pwdtxt.Trim().Length == 0 )
			{
				Console.WriteLine( "Password is empty. Please try again. Press enter to close." );
				return;
			}
			SecureString pwd = new SecureString();
			foreach ( char c in pwdtxt.ToCharArray() )
				pwd.AppendChar( c );
			#endregion

			int scProcessed = 0;

			if ( removeFilteredUsers || validateSitesToProcess )
			{
				#region If Removing Permissions
				SharePointOnlineCredentials creds = new SharePointOnlineCredentials( adminUser, pwd );
				List<string> sitesList = new List<string>();
				try
				{
					using ( var reader = new StreamReader( fileWithSiteUrlsToProcess ) )
					{
						while ( !reader.EndOfStream )
						{
							sitesList.Add( reader.ReadLine().Trim() );
						}
					}
				} catch ( Exception e1 )
				{
					Console.WriteLine( "Error reading csv files contains site urls to process. Error details {0}", e1.ToString() );
					Console.Read();
					return;
				}
				foreach ( string siteurl in sitesList )
				{
					scProcessed++;

					try
					{
						using ( var context = new ClientContext( siteurl ) )
						{
							context.Credentials = creds;
							var rootWeb = context.Web;
							var subWebs = rootWeb.Webs;
							context.Load( rootWeb );
							context.Load( subWebs, w => w.Include( a => a.HasUniqueRoleAssignments, a => a.Url ) );
							context.ExecuteQuery();

							Console.WriteLine( "\n" + scProcessed + ") Site Collection:\t" + siteurl + "\t(" + subWebs.Count + ")" );

							ProcessWeb( rootWeb, creds, true );
							GetAllWebs( subWebs, creds );
						}
					} catch ( Exception e1 )
					{
						Console.WriteLine( "Exception while processing site collection {0}. Error details: {1}", siteurl, e1.ToString() );
					}
				}
				#endregion
			} else
			{
				#region If Querying for Users
				ClientContext adminCtx = new ClientContext( adminSiteUrl );
				SharePointOnlineCredentials creds = new SharePointOnlineCredentials( adminUser, pwd );
				adminCtx.Credentials = creds;
				Tenant tenant = new Tenant( adminCtx );

				int startIndex = 0;
				SPOSitePropertiesEnumerable props = null;

				while ( props == null || props.Count > 0 )
				{

					props = tenant.GetSiteProperties( startIndex, true );
					adminCtx.Load( props );
					adminCtx.ExecuteQuery();
					if ( props != null && props.Count > 0 )
					{
						foreach ( SiteProperties prop in props )
						{
							scProcessed++;
							if ( prop.Url != "https://my.metlife.com/" )
							{
								try
								{
									using ( var context = new ClientContext( prop.Url ) )
									{
										context.Credentials = creds;
										var rootWeb = context.Web;
										var subWebs = rootWeb.Webs;
										context.Load( rootWeb );
										context.Load( subWebs, w => w.Include( a => a.HasUniqueRoleAssignments, a => a.Url ) );
										context.ExecuteQuery();

										Console.WriteLine( "\n" + scProcessed + ") Site Collection:\t" + prop.Url + "\t(" + subWebs.Count + ")" );

										ProcessWeb( rootWeb, creds, true );
										GetAllWebs( subWebs, creds );
									}
								} catch ( Exception e1 )
								{
									Console.WriteLine( "\n" + scProcessed + ") Site Collection EXCEPTION:\t {0}. Error details: {1}", prop.Url, e1.Message );
									txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", prop.Title, prop.Url, prop.Title, "n/a", "Web", "n/a", "n/a", "n/a", "ERROR( " + e1.Message + " )", Environment.NewLine );
									System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
								}
							} else
							{
								Console.WriteLine( "\n" + scProcessed + ") Site Collection:\t (SKIPPING)" + prop.Url );
								txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", prop.Title, prop.Url, prop.Title, "n/a", "Web", "n/a", "n/a", "n/a", "SKIPPED", Environment.NewLine );
								System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
							}
						}
						startIndex += props.Count;
					}
				}
				#endregion
			}

			/** DELETED **/
			Console.WriteLine( "Finished processing all " + scProcessed + " site collections." );
			Console.Read();
		}
		public static void GetAllWebs( WebCollection webs, SharePointOnlineCredentials creds )
		{
			foreach ( var web in webs )
			{
				try
				{
					ProcessWeb( web, creds, false );
					using ( var context = new ClientContext( web.Url ) )
					{
						context.Credentials = creds;
						var subwebs = context.Web.Webs;
						context.Load( subwebs, w => w.Include( a => a.HasUniqueRoleAssignments, a => a.Url ) );
						context.ExecuteQuery();
						if ( subwebs.Count > 0 )
							GetAllWebs( subwebs, creds );
					}
				} catch ( Exception e1 )
				{
					Console.WriteLine( "Exception while processing web {0}. Error details: {1}", web.Url, e1.ToString() );
				}
			}
		}
		public static void ProcessWeb( Web subweb, SharePointOnlineCredentials creds, bool isRootWeb )
		{
			using ( var context = new ClientContext( subweb.Url ) )
			{
				context.Credentials = creds;
				var web = context.Web;
				context.Load( web, w => w.Title, w => w.Url, w => w.HasUniqueRoleAssignments, w => w.RoleAssignments.Include( roleAsg => roleAsg.Member.LoginName,
														roleAsg => roleAsg.Member.Id,
														roleAsg => roleAsg.Member.PrincipalType,
														roleAsg => roleAsg.RoleDefinitionBindings.Include( roleDef => roleDef.Name,
														roleDef => roleDef.Description ) ) );
				context.ExecuteQuery();

				if ( web.Title == "Webtrends Advanced Analytics" )
				{
					Console.WriteLine( "\tWeb: {0} ({1}),  IsRootWeb: {2} (SKIPPED)", web.Url, web.Title, isRootWeb );
				} else
				{
					Console.WriteLine( "\tWeb: {0} ({1}),  IsRootWeb: {2}", web.Url, web.Title, isRootWeb );

					if ( isRootWeb )
					{
						if ( siteGroupIDs != null )
							siteGroupIDs.Clear();
						else
							siteGroupIDs = new List<int>();
					}
					if ( web.HasUniqueRoleAssignments )
						ProcessUniquePerms( context, web, isRootWeb );

					#region Process List Permissions
					if ( !skipLists )
					{
						context.Load( web.Lists, a => a.IncludeWithDefaultProperties( b => b.HasUniqueRoleAssignments ),
								permsn => permsn.Include( a => a.Title, a => a.Id, a => a.RootFolder, a => a.RoleAssignments.Include( roleAsg => roleAsg.Member.LoginName,
															 roleAsg => roleAsg.Member.Id,
															 roleAsg => roleAsg.Member.PrincipalType,
															 roleAsg => roleAsg.RoleDefinitionBindings.Include( roleDef => roleDef.Name,
															 roleDef => roleDef.Description ) ) ) );
						context.ExecuteQuery();
						foreach ( List list in web.Lists )
						{
							try
							{
								if ( !ignoreLists.Contains( list.Title, StringComparer.CurrentCultureIgnoreCase ) && !list.Hidden )
								{
									Console.WriteLine( "\t\tList title:{0}", list.Title );
									//List list = web.Lists.GetByTitle("Documents");

									if ( list.HasUniqueRoleAssignments )
									{
										Console.WriteLine( "\t\t\t{0} has unique perms", list.Title );
										ProcessUniquePerms( context, list );
									}
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
										listItems = list.GetItems( camlQuery );
										context.Load( listItems, l => l.ListItemCollectionPosition, a => a.IncludeWithDefaultProperties( b => b.HasUniqueRoleAssignments, b => b.FileSystemObjectType ),
												permsn => permsn.Include( a => a.RoleAssignments.Include( roleAsg => roleAsg.Member.LoginName,
																 roleAsg => roleAsg.Member.Id,
																 roleAsg => roleAsg.Member.PrincipalType,
																 roleAsg => roleAsg.RoleDefinitionBindings.Include( roleDef => roleDef.Name,
																 roleDef => roleDef.Description ) ) ) );
										context.ExecuteQuery();
										position = listItems.ListItemCollectionPosition;
										allItems.AddRange( listItems.Where( li => li.HasUniqueRoleAssignments == true ) );
										i++;
									}
									while ( position != null );
									//if (i > 0)
									//  Console.WriteLine("Iteration count:" + i);
									Console.WriteLine( "\t\t\tFound {0} unique permissioned list items,files,folders", allItems.Count );
									//if ( listItems != null )
									if ( allItems != null )
									{
										//foreach ( var item in listItems )
										foreach ( var item in allItems )
										{
											ProcessUniquePerms( context, item, list );

											//Moved the logic to check uniquepersm to AddRange method
											//if (item.HasUniqueRoleAssignments)
											//{
											//    ProcessUniquePerms(context, item, list);
											//}
											//else
											//{
											//    //Console.WriteLine("No unique permission found");
											//}
											//Console.WriteLine("###############");
										}
									}
								}
							} catch ( Microsoft.SharePoint.Client.ServerUnauthorizedAccessException noAccessEx )
							{
								Console.WriteLine( "No access to list: " + list.Title );
							} catch ( Exception ex1 )
							{
								Console.WriteLine( "Unknown error processing list: " + list.Title + ". Error details: " + ex1.ToString() );
							}
							//Console.WriteLine("-------------------------------------");
						}
					}
					#endregion
				}
			}
		}
		public static string ProcessLoginName( string memberLoginName )
		{
			if ( memberLoginName.Equals( "c:0(.s|true" ) )
			{
				return "Everyone";
			} else if ( memberLoginName.Contains( "spo-grid-all-users" ) )
			{
				return "Everyone except external users";
			}
			return memberLoginName;
		}
		public static void ProcessUniquePerms( ClientContext context, ListItem item, List list )
		{
			string txt = string.Empty;
			List<RoleAssignment> roleAsgs = new List<RoleAssignment>();
			foreach ( var roleAsg in item.RoleAssignments )
			{
				//Console.WriteLine("User/Group: " + roleAsg.Member.LoginName);
				List<string> roles = new List<string>();
				foreach ( var role in roleAsg.RoleDefinitionBindings )
				{
					roles.Add( role.Name );
				}
				//Console.WriteLine("  Permissions: " + string.Join(",", roles.ToArray()));
				string perms = "\"" + string.Join( "|", roles.ToArray() ) + "\"";
				string pId = ProcessLoginName( roleAsg.Member.LoginName );
				if ( filterUsers.Contains( pId, StringComparer.CurrentCultureIgnoreCase ) )
				{
					if ( removeFilteredUsers )
					{
						roleAsgs.Add( roleAsg );
					} else
					{
						txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, pId, string.Join( "|", roles.ToArray() ), "Granted directly", Environment.NewLine );
						System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
					}
				}
				if ( roleAsg.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup )
				{
					//Process only if the web has unique SPGroup i.e do not process groups inheirted from Rootweb
					if ( !siteGroupIDs.Contains( roleAsg.Member.Id ) )
					{
						List<User> usersToRemove = new List<User>();
						var grp = context.Web.SiteGroups.GetByName( roleAsg.Member.LoginName );
						context.Load( grp, g => g.Users );
						context.ExecuteQuery();
						//Console.WriteLine("  User count: " + grp.Users.Count);
						foreach ( var user in grp.Users )
						{
							pId = string.Empty;
							pId = ProcessLoginName( user.LoginName );
							if ( filterUsers.Contains( pId, StringComparer.CurrentCultureIgnoreCase ) )
							{
								if ( removeFilteredUsers )
								{
									usersToRemove.Add( user );
								} else
								{
									txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, pId, string.Join( "|", roles.ToArray() ), "Granted  Through Group Membership. Group name:" + roleAsg.Member.LoginName, Environment.NewLine );
									System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
								}
							}
						}
						if ( usersToRemove.Count > 0 )
						{
							foreach ( var user in usersToRemove )
							{
								try
								{
									if ( !WhatIf )
									{
										grp.Users.Remove( user );
										grp.Update();
										context.ExecuteQuery();
									}
									txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, user.LoginName, "", ((WhatIf) ? "WhatIf: " : "") + "User removed from group: " + roleAsg.Member.LoginName, Environment.NewLine );
									System.IO.File.AppendAllText( removeUsersLogFilePath, txt, ASCIIEncoding.UTF8 );
								} catch ( Exception e2 )
								{
									Console.WriteLine( "Failed to remove user {0} from group {1}. Error details: {2}", user.LoginName, grp.LoginName, e2.ToString() );
								}
							}
						}
					}
				}
				//Console.WriteLine(Environment.NewLine);
			}
			if ( roleAsgs.Count > 0 )
			{
				try
				{
					foreach ( var roleAsg in roleAsgs )
					{
						if ( !WhatIf )
						{
							roleAsg.RoleDefinitionBindings.RemoveAll();
							roleAsg.DeleteObject();
							context.ExecuteQuery();
						}
						txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, item["FileRef"], item.FileSystemObjectType, roleAsg.Member.PrincipalType, ProcessLoginName( roleAsg.Member.LoginName ), "", ((WhatIf) ? "WhatIf: " : "") + "User/Group removed", Environment.NewLine );
						System.IO.File.AppendAllText( removeUsersLogFilePath, txt, ASCIIEncoding.UTF8 );
					}
				} catch ( Exception e1 )
				{
					Console.WriteLine( "Failed to remove user/group from list item: {0}. Error details: {1}", item["FileRef"], e1.ToString() );
				}
			}
		}
		public static void ProcessUniquePerms( ClientContext context, List list )
		{
			string txt = string.Empty;
			List<RoleAssignment> roleAsgs = new List<RoleAssignment>();
			foreach ( var roleAsg in list.RoleAssignments )
			{
				//Console.WriteLine("User/Group: " + roleAsg.Member.LoginName);
				List<string> roles = new List<string>();
				foreach ( var role in roleAsg.RoleDefinitionBindings )
				{
					roles.Add( role.Name );
				}
				//Console.WriteLine("  Permissions: " + string.Join(",", roles.ToArray()));
				string perms = "\"" + string.Join( "|", roles.ToArray() ) + "\"";
				string pId = ProcessLoginName( roleAsg.Member.LoginName );
				if ( filterUsers.Contains( pId, StringComparer.CurrentCultureIgnoreCase ) )
				{
					if ( removeFilteredUsers )
					{
						roleAsgs.Add( roleAsg );
					} else
					{
						txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, pId, string.Join( "|", roles.ToArray() ), "Granted directly", Environment.NewLine );
						System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
					}
				}
				if ( roleAsg.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup )
				{
					//Process only if the list has unique SPGroup i.e do not process groups inheirted from Rootweb
					if ( !siteGroupIDs.Contains( roleAsg.Member.Id ) )
					{
						List<User> usersToRemove = new List<User>();
						var grp = context.Web.SiteGroups.GetByName( roleAsg.Member.LoginName );
						context.Load( grp, g => g.Users );
						context.ExecuteQuery();
						//Console.WriteLine("  User count: " + grp.Users.Count);
						foreach ( var user in grp.Users )
						{
							pId = string.Empty;
							pId = ProcessLoginName( user.LoginName );
							if ( filterUsers.Contains( pId, StringComparer.CurrentCultureIgnoreCase ) )
							{
								if ( removeFilteredUsers )
								{
									usersToRemove.Add( user );
								} else
								{
									txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, pId, string.Join( "|", roles.ToArray() ), "Granted  Through Group Membership. Group name:" + roleAsg.Member.LoginName, Environment.NewLine );
									System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
								}
							}
						}
						if ( usersToRemove.Count > 0 )
						{
							foreach ( var user in usersToRemove )
							{
								try
								{
									if ( !WhatIf )
									{
										grp.Users.Remove( user );
										grp.Update();
										context.ExecuteQuery();
									}
									txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, user.LoginName, "", ((WhatIf) ? "WhatIf: " : "") + "User removed from group: " + roleAsg.Member.LoginName, Environment.NewLine );
									System.IO.File.AppendAllText( removeUsersLogFilePath, txt, ASCIIEncoding.UTF8 );
								} catch ( Exception e2 )
								{
									Console.WriteLine( "Failed to remove user {0} from group {1}. Error details: {2}", user.LoginName, grp.LoginName, e2.ToString() );
								}
							}
						}
					}
				}
				//Console.WriteLine(Environment.NewLine);
			}
			if ( roleAsgs.Count > 0 )
			{
				try
				{
					foreach ( var roleAsg in roleAsgs )
					{
						if ( !WhatIf )
						{
							roleAsg.RoleDefinitionBindings.RemoveAll();
							roleAsg.DeleteObject();
							context.ExecuteQuery();
						}
						txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, list.Title, list.RootFolder.ServerRelativeUrl, list.BaseType, roleAsg.Member.PrincipalType, ProcessLoginName( roleAsg.Member.LoginName ), "", ((WhatIf) ? "WhatIf: " : "") + "User/Group removed", Environment.NewLine );
						System.IO.File.AppendAllText( removeUsersLogFilePath, txt, ASCIIEncoding.UTF8 );
					}
				} catch ( Exception e1 )
				{
					Console.WriteLine( "Failed to remove user/group from list: {0}. Error details: {1}", list.Title, e1.ToString() );
				}
			}
		}
		public static void ProcessUniquePerms( ClientContext context, Web web, bool isRootWeb )
		{
			string txt = string.Empty;
			List<RoleAssignment> roleAsgs = new List<RoleAssignment>();
			//foreach (var roleAsg in web.RoleAssignments) ;
			int totalRolesAssignments = web.RoleAssignments.Count;
			for ( var i = 0; i < totalRolesAssignments; i++ )
			{
				var roleAsg = web.RoleAssignments[i];
				//Console.WriteLine("User/Group: " + roleAsg.Member.LoginName);
				List<string> roles = new List<string>();
				foreach ( var role in roleAsg.RoleDefinitionBindings )
				{
					roles.Add( role.Name );
				}
				//Console.WriteLine("  Permissions: " + string.Join(",", roles.ToArray()));
				string perms = "\"" + string.Join( "|", roles.ToArray() ) + "\"";
				string pId = ProcessLoginName( roleAsg.Member.LoginName );
				if ( filterUsers.Contains( pId, StringComparer.CurrentCultureIgnoreCase ) )
				{
					if ( removeFilteredUsers )
					{
						roleAsgs.Add( roleAsg );
					} else
					{
						txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, web.Title, "n/a", "Web", roleAsg.Member.PrincipalType, pId, string.Join( "|", roles.ToArray() ), "Granted directly", Environment.NewLine );
						System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
					}
				}
				if ( roleAsg.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup )
				{
					//Process only if the web has unique SPGroup i.e do not process groups inheirted from Rootweb
					if ( !siteGroupIDs.Contains( roleAsg.Member.Id ) )
					{
						List<User> usersToRemove = new List<User>();
						var grp = context.Web.SiteGroups.GetByName( roleAsg.Member.LoginName );
						context.Load( grp, g => g.Users, g => g.Id );
						context.ExecuteQuery();
						if ( isRootWeb ) siteGroupIDs.Add( grp.Id );
						//Console.WriteLine("Groupname: {0}, ID: {1}, RoleAsg_ID: {2}, User count: {3}", roleAsg.Member.LoginName, grp.Id, roleAsg.Member.Id, grp.Users.Count);
						foreach ( var user in grp.Users )
						{
							pId = string.Empty;
							pId = ProcessLoginName( user.LoginName );
							if ( filterUsers.Contains( pId, StringComparer.CurrentCultureIgnoreCase ) )
							{
								if ( removeFilteredUsers )
								{
									usersToRemove.Add( user );
								} else
								{
									txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, web.Title, "n/a", "Web", roleAsg.Member.PrincipalType, pId, string.Join( "|", roles.ToArray() ), "Granted  Through Group Membership. Group name:" + roleAsg.Member.LoginName, Environment.NewLine );
									System.IO.File.AppendAllText( filePath, txt, ASCIIEncoding.UTF8 );
								}


							}
						}
						if ( usersToRemove.Count > 0 )
						{
							foreach ( var user in usersToRemove )
							{
								try
								{
									if ( !WhatIf )
									{
										grp.Users.Remove( user );
										grp.Update();
										context.ExecuteQuery();
									}
									txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, "Removed user/group", "n/a", "Web", roleAsg.Member.PrincipalType, user.LoginName, "", ((WhatIf) ? "WhatIf: " : "") + "User removed from group: " + roleAsg.Member.LoginName, Environment.NewLine );
									System.IO.File.AppendAllText( removeUsersLogFilePath, txt, ASCIIEncoding.UTF8 );
								} catch ( Exception e2 )
								{
									Console.WriteLine( "Failed to remove user {0} from group {1}. Error details: {2}", user.LoginName, grp.LoginName, e2.ToString() );
								}
							}
						}
					}
				}
				//Console.WriteLine(Environment.NewLine);
			}

			if ( roleAsgs.Count > 0 )
			{
				try
				{
					foreach ( var roleAsg in roleAsgs )
					{
						if ( !WhatIf )
						{
							roleAsg.RoleDefinitionBindings.RemoveAll();
							roleAsg.DeleteObject();
							context.ExecuteQuery();
						}
						txt = string.Format( "\"{0}\", {1}, \"{2}\", {3}, {4}, {5}, {6}, {7}, \"{8}\"{9}", context.Web.Title, context.Web.Url, "Removed user/group", "n/a", "Web", roleAsg.Member.PrincipalType, ProcessLoginName( roleAsg.Member.LoginName ), "", ((WhatIf) ? "WhatIf: " : "") + "User/Group removed", Environment.NewLine );
						System.IO.File.AppendAllText( removeUsersLogFilePath, txt, ASCIIEncoding.UTF8 );
					}
				} catch ( Exception e1 )
				{
					Console.WriteLine( "Failed to remove user/group from site: {0}. Error details: {1}", web.Url, e1.ToString() );
				}
			}
		}
	}
}
