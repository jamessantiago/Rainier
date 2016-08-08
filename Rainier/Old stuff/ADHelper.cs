using System;
using System.Collections.Generic;
using System.Linq;

using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;

using System.Collections;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using ActiveDs;
using System.Configuration;

using ADManager.Models;

namespace ADManager.Util
{
    public static class ADHelper
    {
        public static string GetServiceAccount()
        {
            if (string.IsNullOrEmpty(serviceAccount))
            {
                serviceAccount = ConfigurationManager.AppSettings["serviceAccount"];
            }
            return serviceAccount;
        }

        public static string GetServicePassword()
        {
            if (string.IsNullOrEmpty(servicePassword))
            {
                servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            }
            return servicePassword;
        }

        public static string serviceAccount = "";
        public static string servicePassword = "";
        //public static string domainController = "ckwlkera21sig81.kor.ds.cmil.mil";
        public static string domainController = "kor.ds.cmil.mil";
		private static string cat = "meow";


        /*
        public static DirectoryEntry GetDirectoryEntry()
        {
            DirectoryEntry de = new DirectoryEntry();
            de.Path = "LDAP://cktang2a2jcis01.cfc.kor.cmil.mil/CN=Users,DC=ds,DC=cmil,DC=mil";
            de.AuthenticationType = AuthenticationTypes.Secure;
            de.Username = ADHelper.serviceAccount;
            de.Password = ADHelper.GetServicePassword();
            // "LDAP://192.168.0.2/cn=eugene park,ou=Users,ou=CJ64,ou=CJ6,ou=Default.Users,DC=ds,DC=cmil,DC=mil"
            // de.Username = "KOR\\eugene.park";
            // de.Password = "1qaz2!QAZ@";
            return de;
        }
         */
        /*
        public static DirectoryEntry GetDirectoryEntry(string path)
        {
            // var fullPath = "LDAP://cktang2a2jcis01.cfc.kor.cmil.mil/" + path + ",DC=ds,DC=cmil,DC=mil";
            var fullPath = "LDAP://" + domainController + "/" + path;
            DirectoryEntry de = new DirectoryEntry(fullPath);
            de.AuthenticationType = AuthenticationTypes.Secure;
            de.Username = ADHelper.serviceAccount;
            de.Password = ADHelper.GetServicePassword();
            return de;
        }
        */
        public static DirectoryEntry GetDirectoryEntryAlt(string path)
        {
            DirectoryEntry de = new DirectoryEntry(path);
            de.AuthenticationType = AuthenticationTypes.Secure;
            de.Username = ADHelper.GetServiceAccount();
            de.Password = ADHelper.GetServicePassword();
            return de;
        }
        public static DirectoryEntry GetDirectoryEntryFull(string path)
        {
            DirectoryEntry de = new DirectoryEntry();
            de.Path = "LDAP://" + domainController + "/" + path;
            de.AuthenticationType = AuthenticationTypes.Secure;
            de.Username = ADHelper.GetServiceAccount();
            de.Password = ADHelper.GetServicePassword();
            return de;
        }
        public static DirectoryEntry GetDirectoryEntryRoot()
        {
            DirectoryEntry de = new DirectoryEntry();
            de.Path = "LDAP://" + domainController + "/DC=kor,DC=ds,DC=cmil,DC=mil";
            de.AuthenticationType = AuthenticationTypes.Secure;
            de.Username = ADHelper.GetServiceAccount();
            de.Password = ADHelper.GetServicePassword();
            return de;
        }
        public static DirectoryEntry GetDirectoryEntryWinNT(string user)
        {
            DirectoryEntry de = new DirectoryEntry("WinNT://KOR/" + user, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
            return de;
        }
        public static DirectoryEntry GetDirectoryEntryWinNT2(string user)
        {
            DirectoryEntry de = new DirectoryEntry("WinNT://KOR/" + user);
            return de;
        }

        public static void SetProperty(DirectoryEntry de, string PropertyName, string PropertyValue)
        {
            //try
            //{
            if (PropertyValue != null)
            {
                if (de.Properties.Contains(PropertyName))
                {
                    de.Properties[PropertyName][0] = PropertyValue;
                }
                else
                {
                    de.Properties[PropertyName].Value = new string[] { PropertyValue };
                }
            }
            //}
            //catch (System.Runtime.InteropServices.COMException e)
            //{
            //    Console.WriteLine(e.ToString());
            //}
        }

        public static string GetProperty(DirectoryEntry de, string PropertyName)
        {
			try
			{
				de.AuthenticationType = de.AuthenticationType = AuthenticationTypes.Secure;
				de.Username = ADHelper.GetServiceAccount();
				de.Password = ADHelper.GetServicePassword();
				if (de.Properties.Contains(PropertyName) && de.Properties[PropertyName] != null && de.Properties[PropertyName].Count > 0)
				{
					return (string)de.Properties[PropertyName][0];
				}
				else return "";
			}
			catch (Exception e)
			{
				EMailUtil.EmailMike("ADHelper.GetProperty", PropertyName);
				return "";
			}
        }

        public static List<string> GetPropertyList(DirectoryEntry de, string PropertyName)
        {
            List<string> p = new List<string>();
            if (de.Properties.Contains(PropertyName) && de.Properties[PropertyName] != null && de.Properties[PropertyName].Count > 0)
            {
                for (int i = 0; i < de.Properties[PropertyName].Count; i++) {
                    p.Add((string)de.Properties[PropertyName][i]);
                }
                return p;
            }
            else return p;
        }

        private static bool UserExists(string UserName)
        {
            DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;
            deSearch.Filter = "(&(objectClass=user) (cn=" + UserName + "))";
            SearchResultCollection results = deSearch.FindAll();
            de.Close();
            if (results.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public static bool EnableAccount(DirectoryEntry de)
        {
			bool answer = false;
            try
            {
				de.AuthenticationType = de.AuthenticationType = AuthenticationTypes.Secure | AuthenticationTypes.Sealing | AuthenticationTypes.Signing | AuthenticationTypes.Delegation;
                de.Username = ADHelper.GetServiceAccount();
                de.Password = ADHelper.GetServicePassword();
                int exp = 0;
                int.TryParse(de.Properties["userAccountControl"].Value.ToString(), out exp);
                if ((exp & 2) > 0)
                {
                    //de.Properties["userAccountControl"].Value = ADUserFlags.NormalAccount;
					de.Properties["userAccountControl"].Value = ADS_USER_FLAG.ADS_UF_NORMAL_ACCOUNT;
                    de.CommitChanges();
                }
				//de.Properties["msExchHideFromAddressLists"].Value = false;
                AddUserToGroup(de, "All-CENTRIXS");
				answer = true;
            }
            catch (Exception e)
            {
				string p = "";
				if (de != null) p = de.Path;
                EMailUtil.EmailMike("EnableAccount(" + p + ")", e.ToString());
            }
			return answer;
        }

        public static void EnableComputerAccount(DirectoryEntry de)
        {
            de.AuthenticationType = AuthenticationTypes.Secure;
            de.Username = ADHelper.GetServiceAccount();
            de.Password = ADHelper.GetServicePassword();
			de.UsePropertyCache = false;
            int exp = (int)de.Properties["userAccountControl"].Value;
            if ((exp & 2) > 0)
            {
                de.Properties["userAccountControl"].Value = 0x1020;
                de.CommitChanges();
            }
            
        }

        public static bool DisableAccount(DirectoryEntry de)
        {
			bool answer = false;
			try
			{
				int exp = (int)de.Properties["userAccountControl"].Value;
				de.Properties["userAccountControl"].Value = 0x0202;
				//de.Properties["msExchHideFromAddressLists"].Value = true;
				de.CommitChanges();
				answer = true;
				EMailUtil.EmailMike("Account Disabled", de.Path);
			}
			catch (Exception e)
			{
				EMailUtil.EmailMike("DisableAccount", e.ToString());
			}
			return answer;
        }

		public static bool BlacklistUser(DirectoryEntry de, string description)
		{
			bool answer = false;
			try
			{
				if (!DisableAccount(de))
					throw new Exception("Account disable failed for " + de.Name);
				string currentDescription = de.Properties["description"].Value.ToString();
				currentDescription = currentDescription.Replace('[', ' ').Replace(']', ' ');
				description = description.Replace('[', ' ').Replace(']', ' ');
				currentDescription = description + " [" + currentDescription + "]";
				SetProperty(de, "description", currentDescription);
				de.CommitChanges();
				answer = true;
			}
			catch (Exception ex)
			{
				EMailUtil.EmailMike("BlacklistUser", ex.ToString());
			}
			return answer;
		}

		public static bool UnBlacklistUser(DirectoryEntry de)
		{
			bool answer = false;
			try
			{
				if (!EnableAccount(de))
					throw new Exception("Account enable failed for " + de.Name);
				string description = de.Properties["description"].Value.ToString();
				if (description.Contains(']') && description.Contains('['))
				{
					description = description.Substring(description.IndexOf('[') + 1, description.IndexOf(']') - description.IndexOf('[') - 1);
					SetProperty(de, "description", description);
					de.CommitChanges();
				}
				answer = true;
			}
			catch (Exception ex)
			{
				DisableAccount(de);
				EMailUtil.EmailMike("UnBlacklistUser", ex.ToString());
			}
			return answer;
		}

        public static string AddUserToGroup(string UserName, string GroupName)
        {
            string error = null;
            DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
            try
            {
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=user) (sAMAccountName=" + UserName + "))";
                ds.SearchScope = SearchScope.Subtree;

                SearchResult results = ds.FindOne();
                if (results != null)
                {
                    DirectoryEntry deUser = ADHelper.GetDirectoryEntryAlt(results.Path);

                    ds.Filter = "(&(objectClass=group) (cn=" + GroupName + "))";
                    SearchResult groupresults = ds.FindOne();
                    if (groupresults != null)
                    {
                        DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(groupresults.Path);
                        //group.AuthenticationType = AuthenticationTypes.Secure;
                        //group.Username = ADHelper.serviceAccount;
                        //group.Username = ADHelper.GetServicePassword();
                        // group.Properties["member"].Add(deUser);
                        group.Invoke("Add", new object[] { deUser.Path.ToString() });
                        group.CommitChanges();
                        group.Close();
                        error = "Successfully Added User " + UserName + " To Group " + GroupName;
                    }
                    else
                    {
                        deUser.Close();
                        error = "Group " + GroupName + " Not Found";
                    }
                }
                else
                {
                    error = "User " + UserName + " Does Not Exist";
                }
            }
            catch (Exception e)
            {
                EMailUtil.EmailMike("Add User To Group", e.ToString());
                return e.ToString();
            }
            finally
            {
                de.Close();
            }
            return error;
        }

		public static string AddUserToDistro(string UserName, string GroupName)
		{
			string error = null;
			DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
			try
			{
				DirectorySearcher ds = new DirectorySearcher();
				ds.SearchRoot = de;
				ds.Filter = "(&(objectClass=user) (sAMAccountName=" + UserName + "))";
				ds.SearchScope = SearchScope.Subtree;

				SearchResult results = ds.FindOne();
				if (results != null)
				{
					DirectoryEntry deUser = ADHelper.GetDirectoryEntryAlt(results.Path);

					ds.Filter = "(&(objectClass=group) (sAMAccountName=" + GroupName + "))";
					SearchResult groupresults = ds.FindOne();
					if (groupresults != null)
					{
						DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(groupresults.Path);
						//group.AuthenticationType = AuthenticationTypes.Secure;
						//group.Username = ADHelper.serviceAccount;
						//group.Username = ADHelper.GetServicePassword();
						// group.Properties["member"].Add(deUser);
						group.Invoke("Add", new object[] { deUser.Path.ToString() });
						group.CommitChanges();
						group.Close();
						error = "Successfully Added User " + UserName + " To Distro " + GroupName;
						EMailUtil.EmailMike("Distro Add", GroupName + ", " + UserName);
					}
					else
					{
						deUser.Close();
						error = "Distro " + GroupName + " Not Found";
					}
				}
				else
				{
					error = "User " + UserName + " Does Not Exist";
				}
			}
			catch (Exception e)
			{
				EMailUtil.EmailMike("Add User To Distro", e.ToString());
				return e.ToString();
			}
			finally
			{
				de.Close();
			}
			return error;
		}

        public static string GetDistroManager(string GroupName)
        {
            try
            {
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=group) (sAMAccountName=" + GroupName + "))";
                ds.SearchScope = SearchScope.Subtree;
                SearchResult results = ds.FindOne();
                DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(results.Path);
                DirectoryEntry user = ADHelper.GetDirectoryEntryAlt("LDAP://" + domainController + "/" + group.Properties["managedBy"].Value.ToString());
                string username = user.Properties["displayName"].Value.ToString();
                if (user != null)
                {
                    IADsSecurityDescriptor sd = (IADsSecurityDescriptor)group.Properties["ntSecurityDescriptor"].Value;
                    IADsAccessControlList acl = (IADsAccessControlList)sd.DiscretionaryAcl;

                    
                        foreach (IADsAccessControlEntry entry in acl)
                        {
                            if (entry.Trustee.Contains(user.Properties["sAMAccountName"].Value.ToString()) && entry.ObjectType == ADACEObjectTypes.ADS_OBJECT_WRITE_MEMBERS)
                            {
                                username += " (Manager can update membership list)";
                            }
                        }
                        if (!username.EndsWith("list)"))
                            username += " (Manager can NOT update membership list, re-select Manager to change this)";
                    
                }                

                return username;
            }
            catch { }
            return "";
        }

        public static string AddManagerToDistro(string UserName, string GroupName)
        {
            string error = null;
            DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
            try
            {
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=user) (sAMAccountName=" + UserName + "))";
                ds.SearchScope = SearchScope.Subtree;

                SearchResult results = ds.FindOne();
                if (results != null)
                {
                    DirectoryEntry deUser = ADHelper.GetDirectoryEntryAlt(results.Path);

                    ds.Filter = "(&(objectClass=group) (sAMAccountName=" + GroupName + "))";
                    SearchResult groupresults = ds.FindOne();
                    if (groupresults != null)
                    {
                        DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(groupresults.Path);
                        group.AuthenticationType = AuthenticationTypes.Secure;
                        group.Username = ADHelper.GetServiceAccount();
                        group.Password = ADHelper.GetServicePassword();

                        DirectoryEntry oldManager;
                        try { 
                            oldManager = GetDirectoryEntryAlt("LDAP://" + domainController + "/" + group.Properties["managedBy"].Value.ToString());
                        }
                        catch{ oldManager = null;}

                        group.ObjectSecurity.SetOwner(new NTAccount(ADHelper.GetServiceAccount()));
                        group.CommitChanges();

                        group.Properties["managedBy"].Value = deUser.Properties["distinguishedName"].Value;
                        group.CommitChanges();

                        IADsSecurityDescriptor sd = (IADsSecurityDescriptor)group.Properties["ntSecurityDescriptor"].Value;
                        IADsAccessControlList acl = (IADsAccessControlList)sd.DiscretionaryAcl;

                        if (oldManager != null)
                        {
                            foreach (IADsAccessControlEntry entry in acl)
                            {
                                if (entry.Trustee.Contains(oldManager.Properties["sAMAccountName"].Value.ToString()) && entry.ObjectType == ADACEObjectTypes.ADS_OBJECT_WRITE_MEMBERS)
                                {
                                    acl.RemoveAce(entry);
                                }
                            }
                        }

                        ActiveDirectoryAccessRule rule = new ActiveDirectoryAccessRule(new NTAccount("kor\\" + deUser.Properties["sAMAccountName"].Value), ActiveDirectoryRights.WriteProperty, System.Security.AccessControl.AccessControlType.Allow, new Guid(ADACEObjectTypes.ADS_OBJECT_WRITE_MEMBERS));

                        group.ObjectSecurity.AddAccessRule(rule);
                        //IADsAccessControlEntry ace_accessControlEntry = new AccessControlEntry();
                        //ace_accessControlEntry.Trustee = "kor\\" + deUser.Properties["sAMAccountName"].Value;
                        //ace_accessControlEntry.AccessMask = (int)ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP;
                        //ace_accessControlEntry.AceFlags = 0;
                        //ace_accessControlEntry.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
                        //ace_accessControlEntry.Flags = (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
                        //ace_accessControlEntry.ObjectType = ADACEObjectTypes.ADS_OBJECT_WRITE_MEMBERS;
                        
                        //acl.AddAce(ace_accessControlEntry);

                        //sd.DiscretionaryAcl = acl;
                        
                        //group.Properties["ntSecurityDescriptor"].Value = sd;
                        
                        group.CommitChanges();

                        group.Close();
                        error = "Successfully Added User " + UserName + " as manager of Distro " + GroupName;
                        EMailUtil.EmailMike("Distro Manager Add", GroupName + ", " + UserName);
                    }
                    else
                    {
                        deUser.Close();
                        error = "Distro " + GroupName + " Not Found";
                    }
                }
                else
                {
                    error = "User " + UserName + " Does Not Exist";
                }
            }
            catch (Exception e)
            {
                EMailUtil.EmailMike("Add User To Distro", e.ToString());
                return e.ToString();
            }
            finally
            {
                de.Close();
            }
            return error;
        }
		
		public static void AddUserToGroup(DirectoryEntry deUser, string GroupName)
        {
            // public static void AddUserToGroup(DirectoryEntry de, DirectoryEntry deUser, string GroupName){
            DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
            try
            {

                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=group) (cn=" + GroupName + "))";
                ds.SearchScope = SearchScope.Subtree;

                SearchResult results = ds.FindOne();
                if (results != null)
                {
                    DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(results.Path);
                    group.AuthenticationType = AuthenticationTypes.Secure;
                    group.Username = ADHelper.serviceAccount;
                    group.Password = ADHelper.GetServicePassword();
                    group.Invoke("Add", new object[] { deUser.Path.ToString() });
                    group.CommitChanges();
                    group.Close();
                }
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            finally
            {
                de.Close();
            }
        }

        public static string RemoveUserFromGroup(string UserName, string GroupName)
        {
            string error = null;
            try
            {
                DirectoryEntry deUser = getUserDE(UserName);
                if (isUserInGroup(deUser, GroupName))
                {
                    DirectoryEntry deGroup = getGroupDE(GroupName);
                    if (deUser != null && deGroup != null)
                    {
                        deGroup.AuthenticationType = AuthenticationTypes.Secure;
                        deGroup.Username = ADHelper.GetServiceAccount();
                        deGroup.Password = ADHelper.GetServicePassword();
                        deGroup.Invoke("Remove", new object[] { deUser.Path.ToString() });
                        deGroup.CommitChanges();
                        deGroup.Close();
                    }
                    deUser.Close();
                }
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
                return e.ToString();
            }
            return error;
        }

		public static string RemoveUserFromDistro(string UserName, string GroupName)
		{
			string error = null;
			try
			{
				DirectoryEntry deUser = getUserDE(UserName);
				try
				{
					DirectoryEntry deGroup = getDistroDE(GroupName);
					if (deUser != null && deGroup != null)
					{
						deGroup.AuthenticationType = AuthenticationTypes.Secure;
						deGroup.Username = ADHelper.GetServiceAccount();
						deGroup.Password = ADHelper.GetServicePassword();
						deGroup.Invoke("Remove", new object[] { deUser.Path.ToString() });
						deGroup.CommitChanges();
						deGroup.Close();
					}
					deUser.Close();
					EMailUtil.EmailMike("Distro Remove", GroupName + ", " + UserName);
				}
				catch (Exception ee)
				{
				}
			}
			catch (Exception e)
			{
				Console.Write(e.ToString());
				return e.ToString();
			}
			return error;
		}

        public static List<ADUser> ListGroupMembers(string GroupName)
        {
            // SortedList groupMembers = new SortedList();
            PropertyValueCollection members = null;
            List<ADUser> groupmemberlist = new List<ADUser>();
            try
            {

                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=group) (SAMAccountName=" + GroupName + "))";
                ds.SearchScope = SearchScope.Subtree;
                //ds.Sort = new SortOption("sAMAccountName", SortDirection.Descending);
                SearchResult groupresult = ds.FindOne();
                if (groupresult != null)
                {
                    DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(groupresult.Path);
                    members = (PropertyValueCollection)group.Properties["member"];
                    foreach (object res in members)
                    {
						try
						{
							string userpath = res.ToString();
							if (!string.IsNullOrEmpty(userpath)) userpath = userpath.Replace("/", "\\/");
							DirectoryEntry deUser = ADHelper.GetDirectoryEntryAlt("LDAP://" + domainController + "/" + userpath);
							ADUser aduser = new ADUser();
							aduser.DisplayName = (string)deUser.Properties["DisplayName"][0];
							aduser.Username = (string)deUser.Properties["sAMAccountName"][0];
							groupmemberlist.Add(aduser);
						}
						catch (Exception einner)
						{
							EMailUtil.EmailMike("DistroMembersDiv", einner.ToString());
						}
                    }
                    group.Close();
                }
                /*
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=group) (cn=" + GroupName + "))";
                ds.SearchScope = SearchScope.Subtree;

                SearchResult groupresult = ds.FindOne();
                if (groupresult != null)
                {
                    DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(groupresult.Path);

                    ds.SearchRoot = de;
                    ds.Filter = "(&(objectCategory=person)(memberOf=CN=" + groupresult.Path + "))";
                    ds.SearchScope = SearchScope.Subtree;

                    ds.PropertiesToLoad.Add("givenName");
                    ds.PropertiesToLoad.Add("samaccountname");
                    ds.PropertiesToLoad.Add("sn");

                    foreach (SearchResult result in ds.FindAll())
                    {
                        groupMembers.Add(result.Properties["samaccountname"][0].ToString(), result.Properties["givenName"][0].ToString());
                        // DirectoryEntry group = ADHelper.GetDirectoryEntry(results[0].Path);
                        // DirectoryEntry group = new DirectoryEntry(results.Path);
                        // DirectoryEntry group = ADHelper.GetDirectoryEntryAlt(results.Path);

                        // group.Properties["member"].Add(deUser);
                        // groupMembers.Add(
                    }
                }      
                 */
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            //groupmemberlist.Sort();
            return groupmemberlist;
        }

        public static bool isUserInGroup(string user, string group)
        {
            bool answer = false;
            DirectoryEntry deUser = getUserDE(user);
            if (deUser != null)
            {
                /*
                IEnumerator x = deUser.Properties.PropertyNames.GetEnumerator();
                string props = "";
                while (x.MoveNext())
                {
                    props += x.Current + ", ";
                }
                */
                foreach (string groups in deUser.Properties["memberOf"])
                {
                    string a = group.ToLower();
                    string b = groups.ToLower();
                    if (b.Contains("cn=" + a))
                    {
                        answer = true;
                    }
                }
                deUser.Close();
            }
            return answer;
        }

        public static bool isUserInGroup(DirectoryEntry deUser, string group)
        {
            bool answer = false;
            if (deUser != null)
            {
                /*
                IEnumerator x = deUser.Properties.PropertyNames.GetEnumerator();
                string props = "";
                while (x.MoveNext())
                {
                    props += x.Current + ", ";
                }
                */
                foreach (string groups in deUser.Properties["memberOf"])
                {
                    string a = group.ToLower();
                    string b = groups.ToLower();
                    if (b.Contains("cn=" + a))
                    {
                        answer = true;
                    }
                }
            }
            return answer;
        }

        public static List<string> GetUserGroups()
        {
            List<string> answer = new List<string>();
            WindowsIdentity blah = WindowsIdentity.GetCurrent();
            foreach (IdentityReference ir in blah.Groups)
            {
                IdentityReference a = ir.Translate(typeof(NTAccount));
                answer.Add(a.Value);
            }
            return answer;
        }

        public static DirectoryEntry getUserDE(string user)
        {
            DirectoryEntry deUser = null;
            try
            {
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=user) (sAMAccountName=" + user + "))";
                ds.SearchScope = SearchScope.Subtree;

                SearchResult results = ds.FindOne();
                if (results != null)
                {
                    deUser = ADHelper.GetDirectoryEntryAlt(results.Path);
                }
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            return deUser;
        }

        public static DirectoryEntry getContactDE(string user)
        {
            DirectoryEntry deUser = null;
            try
            {
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(name=" + user + "))";
                ds.SearchScope = SearchScope.Subtree;

                SearchResult results = ds.FindOne();
                if (results != null)
                {
                    deUser = ADHelper.GetDirectoryEntryAlt(results.Path);
                }
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            return deUser;
        }

        public static DirectoryEntry getGroupDE(string group)
        {
            DirectoryEntry deGroup = null;
            try
            {
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=group) (cn=" + group + "))";
                ds.SearchScope = SearchScope.Subtree;
                //ds.Sort = new SortOption("sAMAccountName", SortDirection.Descending);
                SearchResult groupresult = ds.FindOne();
                if (groupresult != null)
                {
                    deGroup = ADHelper.GetDirectoryEntryAlt(groupresult.Path);
                }
                de.Close();
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            return deGroup;
        }

        public static DirectoryEntry getDistroDE(string group)
        {
            DirectoryEntry deGroup = null;
            try
            {
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectClass=group) (samaccountname=" + group + "))";
                ds.SearchScope = SearchScope.Subtree;
                //ds.Sort = new SortOption("sAMAccountName", SortDirection.Descending);
                SearchResult groupresult = ds.FindOne();
                if (groupresult != null)
                {
                    deGroup = ADHelper.GetDirectoryEntryAlt(groupresult.Path);
                }
                de.Close();
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            return deGroup;
        }

		public static DirectoryEntry getDistroDEAccount(string samaccountname)
		{
			DirectoryEntry deGroup = null;
			try
			{
				DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
				DirectorySearcher ds = new DirectorySearcher();
				ds.SearchRoot = de;
				ds.Filter = "(&(objectClass=group)(SAMAccountName=" + samaccountname + ")(!groupType:1.2.840.113556.1.4.803:=2147483648))";
				ds.SearchScope = SearchScope.Subtree;
				//ds.Sort = new SortOption("sAMAccountName", SortDirection.Descending);
				SearchResult groupresult = ds.FindOne();
				if (groupresult != null)
				{
					deGroup = ADHelper.GetDirectoryEntryAlt(groupresult.Path);
				}
				de.Close();
			}
			catch (Exception e)
			{
				Console.Write(e.ToString());
			}
			return deGroup;
		}

		public static DirectoryEntry getGroupDEAccount(string samaccountname)
		{
			DirectoryEntry deGroup = null;
			try
			{
				DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
				DirectorySearcher ds = new DirectorySearcher();
				ds.SearchRoot = de;
				ds.Filter = "(&(objectClass=group)(SAMAccountName=" + samaccountname + ")(groupType:1.2.840.113556.1.4.803:=2147483648))";
				ds.SearchScope = SearchScope.Subtree;
				//ds.Sort = new SortOption("sAMAccountName", SortDirection.Descending);
				SearchResult groupresult = ds.FindOne();
				if (groupresult != null)
				{
					deGroup = ADHelper.GetDirectoryEntryAlt(groupresult.Path);
				}
				de.Close();
			}
			catch (Exception e)
			{
				Console.Write(e.ToString());
			}
			return deGroup;
		}

        public static ArrayList SearchUsers(string keywords)
        {
            ArrayList searchresults = new ArrayList();
            DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
            try
            {

                DirectorySearcher deSearch = new DirectorySearcher();
                deSearch.SearchRoot = de;
                string searchstring = "";
                if (keywords != null && keywords.Trim().Length > 0)
                {
                    string[] keys = keywords.Split(' ', ',', '.');
                    for (int i = 0; i < keys.Length; i++)
                    {
                        searchstring += "(|(cn=*" + keys[i] + "*)(displayName=*" + keys[i] + "*))";
                    }
                    deSearch.Filter = "(&(objectCategory=Person)(objectClass=user)" + searchstring + ")";
                    SearchResultCollection results = deSearch.FindAll();
                    foreach (SearchResult result in results)
                    {
                        DirectoryEntry deUser = ADHelper.GetDirectoryEntryAlt(result.Path);
                        ADUser aduser = new ADUser();
                        ICollection blah = deUser.Properties.PropertyNames;
                        /*
                        string asdf = "";
                        foreach (string x in blah)
                        {
                            PropertyValueCollection qwer = deUser.Properties[x];
                            asdf += x + "=";
                            IEnumerator qenum = qwer.GetEnumerator();
                            while (qenum.MoveNext())
                            {
                                asdf += qenum.Current + ",";
                            }
                            asdf += "<br />";
                        }
                        */
                        //aduser.DisplayName = asdf;
                        aduser.DisplayName = (string)deUser.Properties["displayName"][0];
                        foreach (string x in blah)
                        {
                            if (x.Equals("name"))
                            {
                                aduser.Username = (string)deUser.Properties["name"][0];
                            }
                        }
                        foreach (string x in blah)
                        {
                            if (x.Equals("sAMAccountName"))
                            {
                                aduser.Username = (string)deUser.Properties["sAMAccountName"][0];
                            }
                        }
                        searchresults.Add(aduser);

                    }
                }
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            finally
            {
                de.Close();
            }
            searchresults.Sort();
            return searchresults;
        }

        public static List<ADUser> GetUsersInOU(DirectoryEntry ou)
        {
            List<ADUser> users = new List<ADUser>();
            DirectorySearcher ds = new DirectorySearcher();
			if (ou != null)
			{
				try
				{
					ds.SearchRoot = ou;
					string filter = "(&(objectCategory=Person)(objectClass=user))";
					ds.Filter = filter;
					ds.Sort.PropertyName = "sn";
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection searchResults = ds.FindAll();
					foreach (SearchResult results in searchResults)
					{
						//DirectoryEntry dey = ADHelper.GetDirectoryEntryAlt(results.Path);
						DirectoryEntry dey = results.GetDirectoryEntry();
						ADUser aduser = GetUserFromDE(dey);
						users.Add(aduser);
						dey.Close();
					}
					ou.Close();
					searchResults.Dispose();
				}
				catch (Exception e)
				{
				}
			}
            return users;
        }

		public static List<ADUser> GetUsersInOUFast(DirectoryEntry ou)
		{
			List<ADUser> users = new List<ADUser>();
			DirectorySearcher ds = new DirectorySearcher();
			if (ou != null)
			{
				try
				{
					ds.SearchRoot = ou;
					string filter = "(&(objectCategory=Person)(objectClass=user))";
					ds.Filter = filter;
					ds.PropertiesToLoad.Add("DisplayName");
					ds.PropertiesToLoad.Add("SAMAccountName");
					ds.Sort.PropertyName = "sn";
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection searchResults = ds.FindAll();
					foreach (SearchResult results in searchResults)
					{
						ADUser aduser = new ADUser();
						aduser.Username = (string)results.Properties["SAMAccountName"][0];
						aduser.DisplayName = (string)results.Properties["DisplayName"][0];
						users.Add(aduser);
					}
					ou.Close();
					searchResults.Dispose();
				}
				catch (Exception e)
				{
				}
			}
			return users;
		}

		public static List<DList> GetDListsInOU(DirectoryEntry ou)
		{
			List<DList> dls = new List<DList>();
			DirectorySearcher ds = new DirectorySearcher();
			if (ou != null)
			{
				try
				{
					ds.SearchRoot = ou;
					string filter = "(&(objectClass=group)(!groupType:1.2.840.113556.1.4.803:=2147483648))";
					ds.Filter = filter;
					ds.Sort.PropertyName = "DisplayName";
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection searchResults = ds.FindAll();
					foreach (SearchResult results in searchResults)
					{
						DirectoryEntry dey = ADHelper.GetDirectoryEntryAlt(results.Path);
						DList dl = GetDListFromDE(dey);
						dls.Add(dl);
						dey.Close();
					}
					ou.Close();
					searchResults.Dispose();
				}
				catch (Exception e)
				{
				}
			}
			return dls;
		}

		public static List<DList> GetGroupsInOU(DirectoryEntry ou)
		{
			List<DList> dls = new List<DList>();
			DirectorySearcher ds = new DirectorySearcher();
			if (ou != null)
			{
				try
				{
					ds.SearchRoot = ou;
					string filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=2147483648))";
					ds.Filter = filter;
					ds.Sort.PropertyName = "DisplayName";
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection searchResults = ds.FindAll();
					foreach (SearchResult results in searchResults)
					{
						DirectoryEntry dey = ADHelper.GetDirectoryEntryAlt(results.Path);
						DList dl = GetDListFromDE(dey);
						dls.Add(dl);
						dey.Close();
					}
					ou.Close();
					searchResults.Dispose();
				}
				catch
				{
				}
			}
			return dls;
		}

		public static string UpdateWorkstation(string WKID, string description)
		{
			DirectorySearcher ds = new DirectorySearcher();
			try
			{
				DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
				ds.SearchRoot = de;
				string filter = "(&(objectClass=computer)(name=" + WKID + "))";
				ds.Filter = filter;
				ds.SearchScope = SearchScope.Subtree;
				ds.PropertiesToLoad.Add("Name");
				ds.PropertyNamesOnly = true;
				SearchResult searchResult = ds.FindOne();
				
				DirectoryEntry dey = ADHelper.GetDirectoryEntryAlt(searchResult.Path);

				dey.AuthenticationType = AuthenticationTypes.Secure;
				dey.Username = ADHelper.GetServiceAccount();
				dey.Password = ADHelper.GetServicePassword();
				ADHelper.SetProperty(dey, "description", description);
				dey.CommitChanges();
				dey.Close();
				return description;
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		public static List<Workstation> GetWorkstationsInOU(DirectoryEntry ou)
		{
			List<Workstation> wks = new List<Workstation>();
			DirectorySearcher ds = new DirectorySearcher();
			if (ou != null)
			{
				try
				{
					ds.SearchRoot = ou;
					string filter = "(&(objectClass=computer))";
					ds.Filter = filter;
					//ds.Sort.PropertyName = "DisplayName";
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection searchResults = ds.FindAll();
					foreach (SearchResult results in searchResults)
					{
						DirectoryEntry dey = ADHelper.GetDirectoryEntryAlt(results.Path);
						Workstation wk = GetWorkstationFromDE(dey);
						wks.Add(wk);
						dey.Close();
					}
					ou.Close();
					searchResults.Dispose();
				}
				catch (Exception e)
				{
				}
			}
			return wks;
		}

        public static bool isUserInOU(DirectoryEntry ou, string username)
        {
            DirectorySearcher ds = new DirectorySearcher();
            ds.SearchRoot = ou;
            ds.Filter = "(&(objectCategory=Person)(objectClass=user)(sAMAccountName=" + username + "))";
            ds.SearchScope = SearchScope.Subtree;
            SearchResult results = ds.FindOne();
            if (results != null) return true;
            else return false;
        }

        public static bool isDistroInOU(DirectoryEntry ou, string distroname)
        {
            DirectorySearcher ds = new DirectorySearcher();
            ds.SearchRoot = ou;
            ds.Filter = "(&(objectClass=group) (sAMAccountName=" + distroname + "))";
            ds.SearchScope = SearchScope.Subtree;
            SearchResult results = ds.FindOne();
            if (results != null) return true;
            else return false;
        }

        public static ADUser GetUserFromDE(DirectoryEntry de)
        {
            ADManagerDataContext _db = new ADManagerDataContext();
            ADUser aduser = new ADUser();
            if (de.Properties["displayName"] != null && de.Properties["displayName"].Count > 0) { aduser.DisplayName = de.Properties["displayName"][0].ToString(); }
            if (de.Properties["givenName"] != null && de.Properties["givenName"].Count > 0) { aduser.FirstName = de.Properties["givenName"][0].ToString(); }
            if (de.Properties["initials"] != null && de.Properties["initials"].Count > 0) { aduser.MI = de.Properties["initials"][0].ToString(); }
            if (de.Properties["sn"] != null && de.Properties["sn"].Count > 0) { aduser.LastName = de.Properties["sn"][0].ToString(); }
            if (de.Properties["sAMAccountName"] != null && de.Properties["sAMAccountName"].Count > 0) { aduser.Username = de.Properties["sAMAccountName"][0].ToString(); }
            if (de.Properties["mail"] != null && de.Properties["mail"].Count > 0) { aduser.Mail = de.Properties["mail"][0].ToString(); }
            if (de.Properties["title"] != null && de.Properties["title"].Count > 0 && de.Properties["title"][0] != null) { aduser.Title = de.Properties["title"][0].ToString(); }
            if (de.Properties["title"] != null && de.Properties["title"].Count > 0 && de.Properties["title"][0] != null) { aduser.Rank = de.Properties["title"][0].ToString(); }
            if (de.Properties["telephoneNumber"] != null && de.Properties["telephoneNumber"].Count > 0 && de.Properties["telephoneNumber"][0] != null) { aduser.Telephone = de.Properties["telephoneNumber"][0].ToString(); }
            if (de.Properties["streetAddress"] != null && de.Properties["streetAddress"].Count > 0 && de.Properties["streetAddress"][0] != null) { aduser.StreetAddress = de.Properties["streetAddress"][0].ToString(); }
            if (de.Properties["l"] != null && de.Properties["l"].Count > 0 && de.Properties["l"][0] != null) { aduser.City = de.Properties["l"][0].ToString(); }
            if (de.Properties["st"] != null && de.Properties["st"].Count > 0 && de.Properties["st"][0] != null) { aduser.State = de.Properties["st"][0].ToString(); }
            if (de.Properties["postalCode"] != null && de.Properties["postalCode"].Count > 0 && de.Properties["postalCode"][0] != null) { aduser.PostalCode = de.Properties["postalCode"][0].ToString(); }
			if (de.Properties["mail"] != null && de.Properties["mail"].Count > 0 && de.Properties["mail"][0] != null) { aduser.Mail = de.Properties["mail"][0].ToString(); }
            if (de.Properties["extensionAttribute5"] != null && de.Properties["extensionAttribute5"].Count > 0 && de.Properties["extensionAttribute5"][0] != null) { aduser.EmailUnclass = de.Properties["extensionAttribute5"][0].ToString(); }
            if (de.Properties["extensionAttribute6"] != null && de.Properties["extensionAttribute6"].Count > 0 && de.Properties["extensionAttribute6"][0] != null) { aduser.EmailClass = de.Properties["extensionAttribute6"][0].ToString(); }
            if (de.Properties["extensionAttribute8"] != null && de.Properties["extensionAttribute8"].Count > 0 && de.Properties["extensionAttribute8"][0] != null) { aduser.Nationality = de.Properties["extensionAttribute8"][0].ToString(); }
            if (de.Properties["extensionAttribute9"] != null && de.Properties["extensionAttribute9"].Count > 0 && de.Properties["extensionAttribute9"][0] != null) { aduser.SSN = de.Properties["extensionAttribute9"][0].ToString(); }
            if (de.Properties["extensionAttribute11"] != null && de.Properties["extensionAttribute11"].Count > 0 && de.Properties["extensionAttribute11"][0] != null) { aduser.Clearance = de.Properties["extensionAttribute11"][0].ToString(); }
            if (de.Properties["extensionAttribute12"] != null && de.Properties["extensionAttribute12"].Count > 0 && de.Properties["extensionAttribute12"][0] != null) { aduser.ClearanceDate = DateTime.Parse(de.Properties["extensionAttribute12"][0].ToString()); }
            if (de.Properties["extensionAttribute13"] != null && de.Properties["extensionAttribute13"].Count > 0 && de.Properties["extensionAttribute13"][0] != null) { aduser.Title = de.Properties["extensionAttribute13"][0].ToString(); }
            if (de.Properties["physicalDeliveryOfficeName"] != null && de.Properties["physicalDeliveryOfficeName"].Count > 0 && de.Properties["physicalDeliveryOfficeName"][0] != null) { aduser.Section = de.Properties["physicalDeliveryOfficeName"][0].ToString(); }
            if (de.Properties["msExchHideFromAddressLists"] != null && de.Properties["msExchHideFromAddressLists"].Count > 0 && de.Properties["msExchHideFromAddressLists"][0] != null) { aduser.IsHideExchange = Convert.ToBoolean(de.Properties["msExchHideFromAddressLists"][0]); }
            if (de.Properties["userAccountControl"] != null && de.Properties["userAccountControl"].Count > 0 && de.Properties["userAccountControl"][0] != null)
            {
                int exp = Convert.ToInt32(de.Properties["userAccountControl"][0]);
                if ((exp & 2) == 0x2)
                {
                    aduser.IsActive = false;
                }
                else aduser.IsActive = true;
            }
            if (de.Properties["LockOutTime"] != null && de.Properties["LockOutTime"].Count > 0 && de.Properties["LockOutTime"][0] != null)
            {
                LargeInteger liLockOut = de.Properties["LockOutTime"].Value as LargeInteger;
                long lt = (((long)(liLockOut.HighPart) << 32) + (long)liLockOut.LowPart);
                if (lt != 0) aduser.IsLocked = true;
                else aduser.IsLocked = false;
            }
            if (de.Properties["accountExpires"] != null && de.Properties["accountExpires"].Count > 0 && de.Properties["accountExpires"][0] != null)
            {
                try
                {
                    LargeInteger liAcctExpiration = de.Properties["accountExpires"].Value as LargeInteger;
                    long dateAcctExpiration = (((long)(liAcctExpiration.HighPart) << 32) + (long)liAcctExpiration.LowPart);
                    aduser.DEROS = DateTime.FromFileTime(dateAcctExpiration);
                }
                catch
                {
                }
            }
            if (de.Properties["description"] != null && de.Properties["description"].Count > 0) { aduser.Unit = de.Properties["description"][0].ToString(); }
            var unitx = _db.Units.SingleOrDefault(p => p.UnitName == aduser.Unit);
            if (unitx != null) aduser.UnitID = unitx.UnitID;
            if (de.Properties["department"] != null && de.Properties["department"].Count > 0) { aduser.Location = de.Properties["department"][0].ToString(); }
            if (aduser.Location == null) aduser.Location = "";
            var locationx = _db.Locations.SingleOrDefault(p => aduser.Location.Equals(p.LocationName));
            if (locationx != null) aduser.LocationID = locationx.LocationID;
            int branchidx = 0;
            try
            {
                if (de.Properties["extensionAttribute7"] != null && de.Properties["extensionAttribute7"].Count > 0)
                {
                    if (Int32.TryParse(de.Properties["extensionAttribute7"][0].ToString(), out branchidx))
                    {
                        aduser.BranchID = branchidx;
                        aduser.Branch = DBHelper.GetBranch(branchidx).BranchName;
                    }
                }
            }
            catch
            {
            }
            aduser.OUPath = de.Parent.Path;
            aduser.OUName = de.Parent.Name;
            try
            {
                OU ou = _db.OUs.SingleOrDefault(x => x.Path.Equals(aduser.OUPath));
                if (ou != null) aduser.OUID = ou.OUID;
            }
            catch { }
            //req.Deros = (DateTime)dey.InvokeGet("AccountExpirationDate");
            de.Close();
            return aduser;
        }

		public static DList GetDListFromDE(DirectoryEntry de)
		{
			DList dl = new DList();
			if (de.Properties["displayName"] != null && de.Properties["displayName"].Count > 0) { dl.Name = de.Properties["displayName"][0].ToString(); }
			if (de.Properties["description"] != null && de.Properties["description"].Count > 0) { dl.Description = de.Properties["description"][0].ToString(); }
			if (de.Properties["sAMAccountName"] != null && de.Properties["sAMAccountName"].Count > 0) { dl.Username = de.Properties["sAMAccountName"][0].ToString(); }
			dl.OUPath = de.Parent.Path;
			dl.OUName = de.Parent.Name;
			try
			{
				ADManagerDataContext _db = new ADManagerDataContext();
				OU ou = _db.OUs.SingleOrDefault(x => x.Path.Equals(dl.OUPath));
				if (ou != null) dl.OUID = ou.OUID;
			}
			catch (Exception ee) { }
			return dl;
		}

		public static DList GetGroupFromDE(DirectoryEntry de)
		{
			DList dl = new DList();
			if (de.Properties["displayName"] != null && de.Properties["displayName"].Count > 0) { dl.Name = de.Properties["displayName"][0].ToString(); }
			if (de.Properties["description"] != null && de.Properties["description"].Count > 0) { dl.Description = de.Properties["description"][0].ToString(); }
			if (de.Properties["sAMAccountName"] != null && de.Properties["sAMAccountName"].Count > 0) { dl.Username = de.Properties["sAMAccountName"][0].ToString(); }
			dl.OUPath = de.Parent.Path;
			dl.OUName = de.Parent.Name;
			try
			{
				ADManagerDataContext _db = new ADManagerDataContext();
				OU ou = _db.OUs.SingleOrDefault(x => x.Path.Equals(dl.OUPath));
				if (ou != null) dl.OUID = ou.OUID;
			}
			catch (Exception ee) { }
			return dl;
		}

		public static Workstation GetWorkstationFromDE(DirectoryEntry de)
		{
			Workstation wk = new Workstation();
			if (de.Properties["cn"] != null && de.Properties["cn"].Count > 0) { wk.Name = de.Properties["cn"][0].ToString(); }
			if (de.Properties["description"] != null && de.Properties["description"].Count > 0) { wk.Description = de.Properties["description"][0].ToString(); }
			if (de.Properties["sAMAccountName"] != null && de.Properties["sAMAccountName"].Count > 0) { wk.Username = de.Properties["sAMAccountName"][0].ToString(); }
			if (de.Properties["userAccountControl"] != null && de.Properties["userAccountControl"].Count > 0 && de.Properties["userAccountControl"][0] != null)
			{
				int exp = Convert.ToInt32(de.Properties["userAccountControl"][0]);
				if ((exp & 2) == 0x2)
				{
					wk.Disabled = true;
				}
				else wk.Disabled = false;
			}
			wk.OUPath = de.Parent.Path;
			wk.OUName = de.Parent.Name;
			try
			{
				ADManagerDataContext _db = new ADManagerDataContext();
				OU ou = _db.OUs.SingleOrDefault(x => x.Path.Equals(wk.OUPath));
				if (ou != null) wk.OUID = ou.OUID;
			}
			catch (Exception ee) { }
			return wk;
		}

		public static List<DirectoryEntry> GetSubDEs(DirectoryEntry de)
		{
			List<DirectoryEntry> answer = new List<DirectoryEntry>();
			answer.Add(de);
			var subs = de.Children.GetEnumerator();
			while (subs.MoveNext())
			{
				DirectoryEntry child = (DirectoryEntry)subs.Current;
				if (child.SchemaClassName == "organizationalUnit") answer = answer.Concat(GetSubDEs(child)).ToList();
			}
			return answer;
		}

        public static ADUser GetUser(string username)
        {
            DirectoryEntry userde = ADHelper.getUserDE(username);
            ADUser user = ADHelper.GetUserFromDE(userde);
            return user;
        }

        public static DirectoryEntry UpdateUserDE(Mod user)
        {
            DirectoryEntry userde = getUserDE(user.Username);
            if (userde != null)
            {
                userde.AuthenticationType = AuthenticationTypes.Secure;
                userde.Username = ADHelper.GetServiceAccount();
                userde.Password = ADHelper.GetServicePassword();
                string displayname = user.LastName.Trim() + ", " + user.FirstName.Trim() + " " + user.Rank.Trim() + " " + Tools.SuperADSafe(user.Unit.Trim()) + " " + user.Nationality.Trim();
				string cn = user.LastName.Trim() + "\\, " + user.FirstName.Trim() + " " + user.Rank.Trim() + " " + Tools.SuperADSafe(user.Unit.Trim()) + " " + user.Nationality.Trim();
                string mailname = user.Username.Trim();
                if (user.BranchID < 5 || user.BranchID == 8)
                {
                    //mailname += "." + user.Rank.Trim().ToLower();
                }
                if (user.BranchID == 6)
                {
                    //mailname += "." + "ctr";
                }
                if ("US".Equals(user.Nationality))
                {
                    mailname += "." + "us";
                }
                else
                {
                    mailname += "." + "ks";
                }
                ADHelper.SetProperty(userde, "givenName", user.FirstName.Trim());
                if (user.MI != null && user.MI.Length > 0) { ADHelper.SetProperty(userde, "initials", user.MI.Trim()); }
                ADHelper.SetProperty(userde, "sn", user.LastName.Trim());
                ADHelper.SetProperty(userde, "DisplayName", displayname.Trim());
                if (user.Telephone != null && user.Telephone.Length > 0)
                {
                    ADHelper.SetProperty(userde, "telephoneNumber", user.Telephone.Trim());
                }
                if (user.Unit != null && user.Unit.Length > 0) ADHelper.SetProperty(userde, "description", user.Unit.Trim());
                if (user.Unit != null && user.Unit.Length > 0) ADHelper.SetProperty(userde, "company", user.Unit.Trim());
                if (user.IPPhone != null && user.IPPhone.Length > 0)
                {
                    ADHelper.SetProperty(userde, "ipPhone", user.IPPhone.Trim());
                    ADHelper.SetProperty(userde, "homePhone", user.IPPhone.Trim());
                }
                userde.CommitChanges();
                if (user.OtherTelephone != null && user.OtherTelephone.Length > 0) ADHelper.SetProperty(userde, "otherTelephone", user.OtherTelephone);
                if (user.MobilePhone != null && user.MobilePhone.Length > 0) ADHelper.SetProperty(userde, "mobile", user.MobilePhone);
                if (user.BranchID > 0) ADHelper.SetProperty(userde, "extensionAttribute7", user.BranchID + "");
                ADHelper.SetProperty(userde, "extensionAttribute8", user.Nationality);
                ADHelper.SetProperty(userde, "extensionAttribute9", user.SSN);
                ADHelper.SetProperty(userde, "department", user.Location);
                ADHelper.SetProperty(userde, "extensionAttribute10", DateTime.Now.ToString());
                ADHelper.SetProperty(userde, "extensionAttribute11", user.Clearance);
                if (user.EmailUnclass != null && user.EmailUnclass.Length > 0) ADHelper.SetProperty(userde, "extensionAttribute5", user.EmailUnclass);
                if (user.EmailClass != null && user.EmailClass.Length > 0) ADHelper.SetProperty(userde, "extensionAttribute6", user.EmailClass);


                // homemdb???
                if (userde.Properties["homeMDB"] != null && userde.Properties["homeMDB"].Count > 0 && userde.Properties["homeMDB"][0] != null)
                {
                    ADHelper.SetProperty(userde, "mailNickName", mailname);
                    ADHelper.SetProperty(userde, "mail", mailname + "@kor.cmil.mil");
                    //DirectoryEntry user = ADHelper.getUserDE(user.Username);
                    ADHelper.SetProperty(userde, "proxyaddresses", "SMTP:" + mailname + "@kor.cmil.mil");
                    //user.CommitChanges();
                    //user.Close();
                }
                userde.CommitChanges();





                if (user.Section != null && user.Section.Trim().Length > 0)
                {
                    ADHelper.SetProperty(userde, "physicalDeliveryOfficeName", user.Section.Trim());
                    //userde.Properties["physicalDeliveryOfficeName"].Add(user.Section.Trim());
                }
                ADHelper.SetProperty(userde, "title", user.Rank.Trim());
                ADHelper.SetProperty(userde, "extensionAttribute13", user.Title.Trim());
                //ADHelper.SetProperty(userde, "extensionAttribute14", user.Augmentee.ToString());
                userde.CommitChanges();
                ADHelper.SetProperty(userde, "streetAddress", user.StreetAddress);
                ADHelper.SetProperty(userde, "l", user.City);
                ADHelper.SetProperty(userde, "st", user.State);
                if (user.PostalCode != null && user.PostalCode.Length > 0) { ADHelper.SetProperty(userde, "postalCode", user.PostalCode); }
                userde.CommitChanges();
                // Set account expiration to DEROS
                if (user.DEROS != null)
                {
                    userde.InvokeSet("AccountExpirationDate", user.DEROS);
                }
                if (user.ClearanceDate != null)
                {
                    ADHelper.SetProperty(userde, "extensionAttribute12", user.ClearanceDate.ToString());
                }
                userde.CommitChanges();
				/*
                if (Convert.ToBoolean(user.IsActive))
                {
                    ADHelper.EnableAccount(userde);
                }
                else ADHelper.DisableAccount(userde);
				*/
                userde.Properties["msExchHideFromAddressLists"].Value = user.IsHideExchange;

                try
                {
                    userde.Rename("CN=" + cn);
                }
                catch (Exception ee)
                {
                    EMailUtil.EmailMike("ADHelper.UpdateUserDE", ee.ToString());
                }
                userde.CommitChanges();
                userde.Close();
            }
            return userde;
        }

        public static DirectoryEntry MoveUserDE(Transfer t)
        {
            DirectoryEntry userde = getUserDE(t.Username);
            try
            {
                if (userde != null)
                {
                    userde.AuthenticationType = AuthenticationTypes.Secure;
                    userde.Username = ADHelper.GetServiceAccount();
                    userde.Password = ADHelper.GetServicePassword();
                    string finalpath = DBHelper.GetFinalPath(t.OUID);
                    DirectoryEntry destination = ADHelper.GetDirectoryEntryAlt(finalpath);
                    userde.MoveTo(destination);
                    userde.CommitChanges();
                    userde.Close();
                }
            }
            catch (Exception e)
            {
                EMailUtil.EmailMike("MoveUserDE", e.ToString());
            }
            return userde;
        }

        public static string RetrieveSafeUsername(Request req)
        {
            string username = "";
            string ln = req.LastName.Trim();
            string fn = req.FirstName.Trim();
            username = fn + "." + ln;
            if (ln.Length > 19)
            {
                username = ln.Substring(0, 19);
            }
            else
            {
                username = ln;
                if (fn.Length + ln.Length > 18)
                {
                    username = fn.Substring(0, 18 - ln.Length) + "." + ln;
                }
                else
                {
                    username = fn + "." + ln;
                }
            }
            string[] tokens = username.Split(' ');
            var tokencounter = 0;
            var tokenizedusername = "";
            foreach (string token in tokens)
            {
                if (tokencounter == 0) tokenizedusername += token;
                else tokenizedusername += "." + token;
                tokencounter++;
            }
            DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
            DirectorySearcher ds = new DirectorySearcher();
            ds.SearchRoot = de;
            string finalusername = tokenizedusername.ToLower();
            int counter = 0;
            bool safe = false;

            while (!safe && counter < 10)
            {
                if (counter > 0) { finalusername = tokenizedusername.ToLower() + counter; }
                ds.Filter = "(&(objectCategory=Person)(objectClass=user)(sAMAccountName=" + finalusername + "))";
                ds.SearchScope = SearchScope.Subtree;
                // SearchResult results = ds.FindOne();
                SearchResultCollection results = ds.FindAll();
                if (results.Count == 0)
                {
                    safe = true;
                }
                counter++;
            }
            de.Close();
            return finalusername;
        }

		public static string RetrieveSafeGroupname(DListRequest req)
		{
			string username = "";
			string n = req.Name.Trim();
			username = n;
			if (n.Length > 19)
			{
				username = n.Substring(0, 19);
			}
			string[] tokens = username.Split(' ');
			var tokencounter = 0;
			var tokenizedusername = "";
			foreach (string token in tokens)
			{
				if (tokencounter == 0) tokenizedusername += token;
				else tokenizedusername += "." + token;
				tokencounter++;
			}
			DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
			DirectorySearcher ds = new DirectorySearcher();
			ds.SearchRoot = de;
			string finalusername = tokenizedusername.ToLower();
			int counter = 0;
			bool safe = false;

			while (!safe && counter < 10)
			{
				if (counter > 0) { finalusername = tokenizedusername.ToLower() + counter; }
				ds.Filter = "(&(objectClass=group)(SAMAccountName=" + finalusername + "))";
				ds.SearchScope = SearchScope.Subtree;
				// SearchResult results = ds.FindOne();
				SearchResultCollection results = ds.FindAll();
				if (results.Count == 0)
				{
					safe = true;
				}
				counter++;
			}
			de.Close();
			return finalusername;
		}

		public static string GetLegacyExchangeDN(string exchangeServer, string mailname)
		{
			string legacyexchangedn = "";
			if (exchangeServer.EndsWith("CKTANG2B9JCIS76") || exchangeServer.EndsWith("CKWLKERB9JCIS76"))
			{
				legacyexchangedn = "/O=KOR/ou=JCISA Administrative Group/cn=Recipients/cn=" + mailname;
			}
			else if (exchangeServer.EndsWith("CKCASEYB92ID071"))
			{
				legacyexchangedn = "/o=KOR/ou=2IDMAIN/cn=Recipients/cn=" + mailname;
			}
			else if (exchangeServer.EndsWith("CUKOR00B91SIG76"))
			{
				legacyexchangedn = "/o=KOR/ou=First Administrative Group/cn=Recipients/cn=" + mailname;
			}
			return legacyexchangedn;
		}

        public static DirectoryEntry CreateUserDE(Request user)
        {
            // Generate display name
            //unit = unit.Replace("#", "\\#");
            string displayname = user.LastName.Trim() + ", " + user.FirstName.Trim() + " " + user.Rank.Trim() + " " + Tools.SuperADSafe(user.Unit.Trim()) + " " + user.Nationality.Trim();
			string cn = user.LastName.Trim() + "\\, " + user.FirstName.Trim() + " " + user.Rank.Trim() + " " + Tools.SuperADSafe(user.Unit.Trim()) + " " + user.Nationality.Trim();
            // String accountdisplayname = lastname + "\\, " + firstname + " " + rank + " " + nationality;
            string mailname = user.Username.Trim();

			

            DirectoryEntry de = ADHelper.GetDirectoryEntryAlt(user.FinalPath);
            de.AuthenticationType = AuthenticationTypes.Secure;
            DirectoryEntries users = de.Children;
            DirectoryEntry newuser = users.Add("CN=" + cn, "user");
            newuser.Username = ADHelper.GetServiceAccount();
            newuser.Password = ADHelper.GetServicePassword();

            try
            {
				if (user.BranchID < 5 || user.BranchID == 8)
				{
					//mailname += "." + rank.Trim().ToLower();
				}
				if (user.BranchID == 6)
				{
					//mailname += "." + "ctr";
				}
				if ("US".Equals(user.Nationality))
				{
					mailname += "." + "us";
				}
				else
				{
					mailname += "." + "ks";
				}

                // Retrieve User's home directory based on last name
                string homedir = user.HomeDirectory;
                string homeMDB = user.HomeMDB;
                string msExchHomeServerName = user.ExchangeServer;
                string homeMTA = user.HomeMTA;
				string legacyExchangeDN = GetLegacyExchangeDN(msExchHomeServerName, mailname);
                newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "sAMAccountName", user.Username.Trim());
                ADHelper.SetProperty(newuser, "givenName", user.FirstName.Trim());
                newuser.CommitChanges();
                if (!string.IsNullOrEmpty(user.MI) && user.MI.Trim().Length > 0) { ADHelper.SetProperty(newuser, "initials", user.MI.Trim()); }
                ADHelper.SetProperty(newuser, "sn", user.LastName.Trim());
                newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "pwdLastSet", "0");
                ADHelper.SetProperty(newuser, "homeDrive", "h:");
                newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "homeDirectory", homedir.Trim());
                ADHelper.SetProperty(newuser, "DisplayName", displayname.Trim());
                ADHelper.SetProperty(newuser, "userPrincipalName", user.Username.Trim() + "@kor.ds.cmil.mil");
                if (!string.IsNullOrEmpty(user.Telephone) && user.Telephone.Trim().Length > 0) ADHelper.SetProperty(newuser, "telephoneNumber", user.Telephone.Trim());
                newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "mailNickName", mailname.Trim());
                ADHelper.SetProperty(newuser, "mail", mailname.Trim() + "@kor.cmil.mil");
				if (!string.IsNullOrEmpty(user.Unit) && user.Unit.Trim().Length > 0) ADHelper.SetProperty(newuser, "description", user.Unit.Trim());
				if (!string.IsNullOrEmpty(user.Unit) && user.Unit.Trim().Length > 0) ADHelper.SetProperty(newuser, "company", user.Unit.Trim());
                newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "scriptPath", "JCISA-Login.bat");
				if (!string.IsNullOrEmpty(user.Section) && user.Section.Trim().Length > 0) ADHelper.SetProperty(newuser, "physicalDeliveryOfficeName", user.Section.Trim());
                ADHelper.SetProperty(newuser, "title", user.Rank.Trim());
                ADHelper.SetProperty(newuser, "extensionAttribute13", user.Title.Trim());
                newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "streetAddress", user.StreetAddress.Trim());
				newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "l", user.City.Trim());
				newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "st", user.State.Trim());
				newuser.CommitChanges();
				if (!string.IsNullOrEmpty(user.PostalCode) && user.PostalCode.Trim().Length > 0) { ADHelper.SetProperty(newuser, "postalCode", user.PostalCode.Trim()); }
                newuser.CommitChanges();
                
				// Mailbox
                ADHelper.SetProperty(newuser, "homeMDB", homeMDB);
                ADHelper.SetProperty(newuser, "msExchHomeServerName", msExchHomeServerName);
                ADHelper.SetProperty(newuser, "homeMTA", homeMTA);
				if (!string.IsNullOrEmpty(legacyExchangeDN)) ADHelper.SetProperty(newuser, "legacyExchangeDN", legacyExchangeDN);
				newuser.CommitChanges();


				if (!string.IsNullOrEmpty(user.IPPhone) && user.IPPhone.Trim().Length > 0)
                {
                    ADHelper.SetProperty(newuser, "ipPhone", user.IPPhone.Trim());
                    ADHelper.SetProperty(newuser, "homePhone", user.IPPhone.Trim());
                }
				if (!string.IsNullOrEmpty(user.OtherTelephone) && user.OtherTelephone.Trim().Length > 0) ADHelper.SetProperty(newuser, "otherTelephone", user.OtherTelephone);
				if (!string.IsNullOrEmpty(user.MobilePhone) && user.MobilePhone.Trim().Length > 0) ADHelper.SetProperty(newuser, "mobile", user.MobilePhone);
                if (user.BranchID > 0) ADHelper.SetProperty(newuser, "extensionAttribute7", user.BranchID + "");
				newuser.CommitChanges();
				ADHelper.SetProperty(newuser, "extensionAttribute8", user.Nationality);
                ADHelper.SetProperty(newuser, "extensionAttribute9", user.SSN);
                ADHelper.SetProperty(newuser, "department", user.Location);
				newuser.CommitChanges();
                ADHelper.SetProperty(newuser, "extensionAttribute10", DateTime.Now.ToString());
                ADHelper.SetProperty(newuser, "extensionAttribute11", user.Clearance);
                ADHelper.SetProperty(newuser, "extensionAttribute12", user.ClearanceDate.ToString());
                //ADHelper.SetProperty(newuser, "extensionAttribute14", augmentee.ToString());
				if (!string.IsNullOrEmpty(user.EmailUnclass) && user.EmailUnclass.Trim().Length > 0) ADHelper.SetProperty(newuser, "extensionAttribute5", user.EmailUnclass);
				if (!string.IsNullOrEmpty(user.EmailClass) && user.EmailClass.Trim().Length > 0) ADHelper.SetProperty(newuser, "extensionAttribute6", user.EmailClass);
                newuser.CommitChanges();

                // 4.  Set account expiration to DEROS
                newuser.InvokeSet("AccountExpirationDate", new object[] { new DateTime(user.DEROS.Year, user.DEROS.Month, user.DEROS.Day) });

                newuser.CommitChanges();
                // AD will not accept enable until user has acceptable password
                //EnableAccount(newuser);

                newuser.Close();
                de.Close();

                // 3. Add user account to groups
				List<long> parentids = DBHelper.GetParentIDs(user.OUID.Value);
				if (parentids.Contains(17))
				{
					ADHelper.AddUserToGroup(newuser, "All-CENTRIXS");
					if (user.LastName.Substring(0, 1).ToLower().CompareTo("9") <= 0)
					{
						ADHelper.AddUserToGroup(newuser, "GG.Users_0-9");
					}
					else if (user.LastName.Substring(0, 1).ToLower().CompareTo("f") <= 0)
					{
						ADHelper.AddUserToGroup(newuser, "GG.Users_A-F");
					}
					else if (user.LastName.Substring(0, 1).ToLower().CompareTo("m") <= 0)
					{
						ADHelper.AddUserToGroup(newuser, "GG.Users_G-M");
					}
					else if (user.LastName.Substring(0, 1).ToLower().CompareTo("r") <= 0)
					{
						ADHelper.AddUserToGroup(newuser, "GG.Users_N-R");
					}
					else if (user.LastName.Substring(0, 1).ToLower().CompareTo("z") <= 0)
					{
						ADHelper.AddUserToGroup(newuser, "GG.Users_S-Z");
					}
				}
				ADHelper.AddUserToGroup(newuser, "GG.DCO.Public");
            }
            catch (Exception e)
            {
                EMailUtil.EmailMike("CreateUserDE - " + user.Username, e.ToString());
				users.Remove(newuser);
				return null;
            }
            return newuser;
        }

        public static DirectoryEntry CreateComputerDE(string path, string computername, string description)
        {
			try
			{
				//path = path.Replace("kor.ds.cmil.mil", "CUWLKERA21SIG81.kor.ds.cmil.mil");
				// Generate display name
				//unit = unit.Replace("#", "\\#");
				string cn = computername;
				DirectoryEntry de = ADHelper.GetDirectoryEntryAlt(path);
				de.AuthenticationType = AuthenticationTypes.Secure;
				de.UsePropertyCache = false;
				DirectoryEntries computers = de.Children;
				DirectoryEntry newcomp = computers.Add("CN=" + cn, "computer");
				newcomp.UsePropertyCache = false;
				newcomp.Username = ADHelper.GetServiceAccount();
				newcomp.Password = ADHelper.GetServicePassword();
				ADHelper.SetProperty(newcomp, "sAMAccountName", computername + "$");
				ADHelper.SetProperty(newcomp, "description", description);
                ADHelper.SetProperty(newcomp, "userAccountControl", "4128");				
				newcomp.CommitChanges();

				

				IADsSecurityDescriptor sd = (IADsSecurityDescriptor)newcomp.Properties["ntSecurityDescriptor"].Value;
				IADsAccessControlList acl = (IADsAccessControlList)sd.DiscretionaryAcl;

				/*
				var acls = acl.GetEnumerator();
				while (acls.MoveNext())
				{
					IADsAccessControlEntry currace = (IADsAccessControlEntry)acls.Current;
					if (currace.Trustee.Equals("KOR\\account.mgt")) acl.RemoveAce(currace);
				}
				*/

				IADsAccessControlEntry ace_writeaccountrestrictions = new AccessControlEntryClass();
				ace_writeaccountrestrictions.Trustee = "Authenticated Users";
				ace_writeaccountrestrictions.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
				ace_writeaccountrestrictions.AceFlags = 0;
				ace_writeaccountrestrictions.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_writeaccountrestrictions.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_writeaccountrestrictions.ObjectType = ADACEObjectTypes.USER_ACCOUNT_RESTRICTIONS;
				IADsAccessControlEntry ace_validatedspn = new AccessControlEntryClass();
				ace_validatedspn.Trustee = "Authenticated Users";
				ace_validatedspn.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_SELF);
				ace_validatedspn.AceFlags = 0;
				ace_validatedspn.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;				
				ace_validatedspn.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_validatedspn.ObjectType = ADACEObjectTypes.VALIDATED_SPN;
				IADsAccessControlEntry ace_validateddns = new AccessControlEntryClass();
				ace_validateddns.Trustee = "Authenticated Users";
				ace_validateddns.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_SELF);
				ace_validateddns.AceFlags = 0;
				ace_validateddns.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_validateddns.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_validateddns.ObjectType = ADACEObjectTypes.VALIDATED_DNS_HOST_NAME;
				IADsAccessControlEntry ace_resetpw = new AccessControlEntryClass();
				ace_resetpw.Trustee = "Authenticated Users";
				ace_resetpw.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
				ace_resetpw.AceFlags = 0;
				ace_resetpw.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_resetpw.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_resetpw.ObjectType = ADACEObjectTypes.RESET_PASSWORD_GUID;
				IADsAccessControlEntry ace_changepw = new AccessControlEntryClass();
				ace_changepw.Trustee = "Authenticated Users";
				ace_changepw.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
				ace_changepw.AceFlags = 0;
				ace_changepw.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_changepw.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_changepw.ObjectType = ADACEObjectTypes.USER_CHANGE_PASSWORD;
				IADsAccessControlEntry ace_receiveas = new AccessControlEntryClass();
				ace_receiveas.Trustee = "Authenticated Users";
				ace_receiveas.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
				ace_receiveas.AceFlags = 0;
				ace_receiveas.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_receiveas.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_receiveas.ObjectType = ADACEObjectTypes.RECEIVE_AS;
				IADsAccessControlEntry ace_sendas = new AccessControlEntryClass();
				ace_sendas.Trustee = "Authenticated Users";
				ace_sendas.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
				ace_sendas.AceFlags = 0;
				ace_sendas.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_sendas.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_sendas.ObjectType = ADACEObjectTypes.SEND_AS;


				IADsAccessControlEntry ace_access = new AccessControlEntryClass();
				ace_access.Trustee = "Authenticated Users";
				ace_access.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS |
												ADS_RIGHTS_ENUM.ADS_RIGHT_ACTRL_DS_LIST |
												ADS_RIGHTS_ENUM.ADS_RIGHT_GENERIC_READ |
												ADS_RIGHTS_ENUM.ADS_RIGHT_DELETE |
												ADS_RIGHTS_ENUM.ADS_RIGHT_DS_DELETE_TREE |
												ADS_RIGHTS_ENUM.ADS_RIGHT_READ_CONTROL);
				ace_access.AceFlags = 0;
				ace_access.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED;


				IADsAccessControlEntry ace_delete = new AccessControlEntryClass();
				ace_delete.Trustee = "Authenticated Users";
				ace_delete.AccessMask = 197076;// (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DELETE | ADS_RIGHTS_ENUM.ADS_RIGHT_DS_DELETE_TREE);
				ace_delete.AceFlags = 0;
				ace_delete.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_delete.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED;
				IADsAccessControlEntry ace_description = new AccessControlEntryClass();
				ace_description.Trustee = "Authenticated Users";
				ace_description.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
				ace_description.AceFlags = 0;
				ace_description.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_description.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_description.ObjectType = ADACEObjectTypes.DESCRIPTION_GUID;
				IADsAccessControlEntry ace_logoninfo = new AccessControlEntryClass();
				ace_logoninfo.Trustee = "Authenticated Users";
				ace_logoninfo.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
				ace_logoninfo.AceFlags = 0;
				ace_logoninfo.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_logoninfo.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_logoninfo.ObjectType = ADACEObjectTypes.USER_LOGON_INFORMATION;
				IADsAccessControlEntry ace_cname = new AccessControlEntryClass();
				ace_cname.Trustee = "Authenticated Users";
				ace_cname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
				ace_cname.AceFlags = 0;
				ace_cname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_cname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_cname.ObjectType = ADACEObjectTypes.COMMONNAME_GUID;
				IADsAccessControlEntry ace_dname = new AccessControlEntryClass();
				ace_dname.Trustee = "Authenticated Users";
				ace_dname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
				ace_dname.AceFlags = 0;
				ace_dname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_dname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_dname.ObjectType = ADACEObjectTypes.DISPLAYNAME_GUID;
				IADsAccessControlEntry ace_sname = new AccessControlEntryClass();
				ace_sname.Trustee = "Authenticated Users";
				ace_sname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
				ace_sname.AceFlags = 0;
				ace_sname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_sname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_sname.ObjectType = ADACEObjectTypes.SAMACCOUNTNAME_GUID;

				IADsAccessControlEntry ace_aaname = new AccessControlEntryClass();
				ace_aaname.Trustee = "Authenticated Users";
				ace_aaname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
				ace_aaname.AceFlags = 0;
				ace_aaname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
				ace_aaname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
				ace_aaname.ObjectType = ADACEObjectTypes.ALLOWED_TO_AUTHENTICATE;

				acl.AddAce(ace_access);
				acl.AddAce(ace_description);
				acl.AddAce(ace_sname);
				acl.AddAce(ace_dname);
				acl.AddAce(ace_logoninfo);
				acl.AddAce(ace_writeaccountrestrictions);
				acl.AddAce(ace_validateddns);
				acl.AddAce(ace_validatedspn);
				acl.AddAce(ace_aaname);
				acl.AddAce(ace_receiveas);
				acl.AddAce(ace_sendas);
				acl.AddAce(ace_changepw);
				acl.AddAce(ace_resetpw);
				acl.AddAce(ace_delete);
				acl.AddAce(ace_cname);


				//IADsAccessControlEntry ace_fullcontrol = new AccessControlEntryClass();
				//ace_fullcontrol.Trustee = "Authenticated Users";
				//ace_fullcontrol.AccessMask = -1;

				//acl.AddAce(ace_fullcontrol);
				
				sd.DiscretionaryAcl = acl;
				//sd.DiscretionaryAcl = ReorderDacl(acl);
				//System.Threading.Thread.Sleep(2000);
				try
				{
					Exception e = null;
					try
					{
						newcomp.Properties["ntSecurityDescriptor"][0] = sd;
						newcomp.CommitChanges();
					}
					catch (Exception ex)
					{
						e = ex;
					}
					if (e != null)
					{
						try
						{

							newcomp.Properties["ntSecurityDescriptor"].Value = sd;
							newcomp.CommitChanges();
							e = null;
						}
						catch (Exception ex)
						{
							e = ex;
						}
					}
					if (e != null)
						throw e;

					

					newcomp.Close();
					de.Close();
					
					EnableComputerAccount(newcomp);

					return newcomp;
				}
				catch (Exception e)
				{
					computers.Remove(newcomp);
					
					EMailUtil.EmailMike("CreateComputerDE(" + path + "," + computername + "," + description + ")", e.ToString());
					return null;
				}
			}
			catch (Exception e)
			{
				EMailUtil.EmailMike("CreateComputerDE(" + path + "," + computername + "," + description + ")", e.ToString());
				
				return null;
			}
        }

		public static object ReorderDacl(object dacl)
		{
			IADsAccessControlList newDacl = new AccessControlList();
			IADsAccessControlList newacl = new AccessControlList();
			IADsAccessControlList impDenyDacl = new AccessControlList();
			IADsAccessControlList inheritedDacl = new AccessControlList();
			IADsAccessControlList impAllowDacl = new AccessControlList();
			IADsAccessControlList imhAllowDacl = new AccessControlList();
			IADsAccessControlList impDenyObjectDacl = new AccessControlList();
			IADsAccessControlList impAllowObjectDacl = new AccessControlList();

			
			foreach (object ace in ((IADsAccessControlList)dacl))
			{
				if ((((AccessControlEntry)ace).AceFlags &
					(int)ADS_ACEFLAG_ENUM.ADS_ACEFLAG_INHERIT_ACE) ==
					(int)ADS_ACEFLAG_ENUM.ADS_ACEFLAG_INHERIT_ACE)
				{
					inheritedDacl.AddAce(ace);					
				}
				else
				{
					switch (((AccessControlEntry)ace).AceFlags)
					{
						case (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED:
							impAllowDacl.AddAce(ace);							
							break;
						case (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_DENIED:
							impDenyDacl.AddAce(ace);
							break;
						case (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT:
							impAllowObjectDacl.AddAce(ace);							
							break;
						case (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_DENIED_OBJECT:
							impDenyObjectDacl.AddAce(ace);							
							break;
						default:
							break;
					}
				}
			}

			foreach (object ace in (AccessControlList)impDenyDacl)
			{
				newacl.AddAce(ace);
			}
			foreach (object ace in (AccessControlList)impDenyObjectDacl)
			{
				newacl.AddAce(ace);
			}
			foreach (object ace in (AccessControlList)impAllowDacl)
			{
				newacl.AddAce(ace);
			}
			foreach (object ace in (AccessControlList)impAllowObjectDacl)
			{
				newacl.AddAce(ace);
			}
			foreach (object ace in (AccessControlList)inheritedDacl)
			{
				newacl.AddAce(ace);
			}

			return newacl;

		}


        public static DirectoryEntry CreateDistroDE(DListRequest req)
        {
            string cn = req.Name;
            DirectoryEntry de = ADHelper.GetDirectoryEntryAlt(req.FinalPath);
            de.AuthenticationType = AuthenticationTypes.Secure;
            DirectoryEntries dls = de.Children;
            DirectoryEntry newdl = dls.Add("CN=" + cn, "group");
            newdl.Username = ADHelper.GetServiceAccount();
            newdl.Password = ADHelper.GetServicePassword();
            DirectoryEntry user = getUserDE(req.RequestBy);
            string blah = user.Path;
            blah = blah.Substring(blah.LastIndexOf("/") + 1);
            ADHelper.SetProperty(newdl, "sAMAccountName", req.Username);
            ADHelper.SetProperty(newdl, "displayName", req.Name);
            ADHelper.SetProperty(newdl, "description", req.Description);
            ADHelper.SetProperty(newdl, "groupType", "2");
            ADHelper.SetProperty(newdl, "mailNickname", req.Username);
            ADHelper.SetProperty(newdl, "mail", req.Username + "@kor.cmil.mil");
            ADHelper.SetProperty(newdl, "LegacyExchangeDN", "/o=KOR/ou=First Administrative Group/cn=Recipients/cn=" + req.Name);
            ADHelper.SetProperty(newdl, "proxyAddresses", "SMTP:" + req.Username + "@kor.cmil.mil");
            ADHelper.SetProperty(newdl, "showInAddressBook", "CN=Default Global Address List,CN=All Global Address Lists,CN=Address Lists Container,CN=KOR,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=ds,DC=cmil,DC=mil");
            ADHelper.SetProperty(newdl, "managedBy", blah);
            //ADHelper.SetProperty(newdl, "reportToOriginator", "True");

            newdl.CommitChanges();
            newdl.Close();
            de.Close();
            return newdl;
        }

        public static DirectoryEntry UpdateDistroManager(string name, string username)
        {
            DirectoryEntry dl = getDistroDE(name);
            try
            {
                if (dl != null)
                {
                    DirectoryEntry user = getUserDE(username);
                    if (user != null)
                    {
                        string blah = user.Path;
                        blah = blah.Substring(blah.LastIndexOf("/") + 1);
                        ADHelper.SetProperty(dl, "managedBy", blah);
                        dl.CommitChanges();
                        dl.Close();
                    }
                }
            }
            catch (Exception e) { e.ToString(); }
            return dl;
        }

        public static SearchResultCollection GetSubOUs(DirectoryEntry rootOU)
        {
            SearchResultCollection results = null;
            try
            {
                DirectorySearcher deSearch = new DirectorySearcher();
                deSearch.SearchRoot = rootOU;
                deSearch.Filter = "(objectClass=organizationalUnit)";
                deSearch.SearchScope = SearchScope.OneLevel;
                results = deSearch.FindAll();
            }
            catch (Exception e)
            {

            }
            return results;
        }
        
        public static List<ADUser> Search(string firstname, string lastname, bool strict)
        {
            List<ADUser> finalResults = new List<ADUser>();
            try
            {
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                string filter = "(&(objectCategory=Person)(objectClass=user)";
                if (strict)
                {
                    if (firstname != null && firstname.Length > 0)
                    {
                        filter += "(givenName=" + firstname + "*)";
                    }
                    if (lastname != null && lastname.Length > 0)
                    {
                        filter += "(sn=" + lastname + "*)";
                    }
                }
                else
                {
                    filter += "(|";
                    if (firstname != null && firstname.Length > 0)
                    {
                        filter += "(givenName=" + firstname + "*)";
                    }
                    if (lastname != null && lastname.Length > 0)
                    {
                        filter += "(sn=" + lastname + "*)";
                    }
                    filter += ")";
                }
                filter += ")";
                // TempData["Message"] = filter;

                ds.Filter = filter;

                // ds.Filter = "(&(objectCategory=Person)(objectClass=user))";
                ds.SearchScope = SearchScope.Subtree;
                SearchResultCollection searchResults = ds.FindAll();

                foreach (SearchResult results in searchResults)
                {
                    DirectoryEntry dey = ADHelper.GetDirectoryEntryAlt(results.Path);
                    ADUser aduser = GetUserFromDE(dey);
                    finalResults.Add(aduser);
                    dey.Close();
                }
                de.Close();
				searchResults.Dispose();
            }
            catch (Exception e)
            {
                return null;
            }
            return finalResults;
        }

        public static bool ResetPassword(string username, string newpassword)
        {
            string error = "";
            try
            {
                DirectoryEntry de = ADHelper.GetDirectoryEntryRoot();
                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                ds.Filter = "(&(objectCategory=Person)(objectClass=user)(sAMAccountName=" + username + "))";
                ds.SearchScope = SearchScope.Subtree;
                SearchResult results = ds.FindOne();
                error += ds.Filter;

                if (results != null)
                {
                    DirectoryEntry newuser = ADHelper.GetDirectoryEntryWinNT(username);
                    newuser.AuthenticationType = AuthenticationTypes.Secure;
                    newuser.Username = ADHelper.GetServiceAccount();
                    newuser.Password = ADHelper.GetServicePassword();
                    object obRet = newuser.Invoke("SetPassword", newpassword);
                    newuser.CommitChanges();
                    newuser.Close();

                    // DirectoryEntry newuser = ADHelper.GetDirectoryEntry(results.Path);
                    // error += " " + newuser.NativeObject;
                    // IADsUser newuserObj = (IADsUser)newuser.NativeObject;

                    // newuserObj.SetPassword("1qaz2!QAZ");
                    // newuserObj.SetInfo();
                    // newuser.InvokeSet("password", "1qaz2!QAZ!");
                    // newuser.Invoke("SetPassword", "1qaz2!QAZ");
                    // newuser.CommitChanges();
                    // production.APasswordReqs.InsertOnSubmit(req);
                    // production.SubmitChanges();
                }

                de.Close();

                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

		//public static string GetCurrUser()
		//{			
		//    string currusername = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().ToLower();
		//    return currusername.Substring(currusername.LastIndexOf('\\') + 1);			
		//}

        public static List<ADUser> GetUsersActive(string iasoname, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.Enabled = true;
                ps.QueryFilter = usr;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersInactive(string iasoname, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.Enabled = false;
                ps.QueryFilter = usr;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersLogon(string iasoname, TimeSpan t, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.AdvancedSearchFilter.LastLogonTime(DateTime.Now.Subtract(t), MatchType.GreaterThan);
                ps.QueryFilter = usr;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersExpire(string iasoname, TimeSpan t, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.Enabled = true;
                usr.AdvancedSearchFilter.AccountExpirationDate(new DateTime(1601, 1, 1), MatchType.NotEquals);
                usr.AdvancedSearchFilter.AccountExpirationDate(DateTime.Now.Add(t), MatchType.LessThan);
                ps.QueryFilter = usr;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersExpired(string iasoname, TimeSpan t, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            //List<OU> roots = DBHelper.GetSearchRoots("jon.ahn");
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.Enabled = true;
                usr.AdvancedSearchFilter.AccountExpirationDate(new DateTime(1601, 1, 1), MatchType.NotEquals);
                usr.AdvancedSearchFilter.AccountExpirationDate(DateTime.Now.Subtract(t), MatchType.GreaterThan);
                usr.AdvancedSearchFilter.AccountExpirationDate(DateTime.Now, MatchType.LessThan);
                ps.QueryFilter = usr;
                
                string temp = (ps.GetUnderlyingSearcher() as DirectorySearcher).Filter;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersExpired1(TimeSpan t, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            //List<OU> roots = DBHelper.GetSearchRoots(GetCurrUser());
            List<OU> roots = DBHelper.GetSearchRoots("jon.ahn");
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                DirectorySearcher deSearch = new DirectorySearcher();
                deSearch.SearchRoot = GetDirectoryEntryAlt(path);
                string searchstring = "(&(accountExpires > " + DateTime.Now.Subtract(t).ToFileTimeUtc() + ")(accountExpires < " + DateTime.Now.ToFileTimeUtc() + ")(!(accountExpires=129283962422814666))(accountExpires=*))";
                deSearch.Filter = "(&(objectCategory=user)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))" + searchstring + ")";
                deSearch.PageSize = 1000;
                deSearch.SearchScope = SearchScope.Subtree;
                SearchResultCollection results = deSearch.FindAll();
                foreach (SearchResult result in results)
                {
                    DirectoryEntry deUser = ADHelper.GetDirectoryEntryAlt(result.Path);
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser((string)deUser.Properties["samAccountName"][0]));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = (string)deUser.Properties["samAccountName"][0];
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersNeverExpire(string iasoname, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.Enabled = true;
                usr.AdvancedSearchFilter.AccountExpirationDate(new DateTime(1601, 1, 1), MatchType.Equals);
                ps.QueryFilter = usr;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersLockedOut(string iasoname, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/')+1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.Enabled = true;
                usr.AdvancedSearchFilter.AccountLockoutTime(new DateTime(1601, 1, 1), MatchType.GreaterThan);
                ps.QueryFilter = usr;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }

        public static List<ADUser> GetUsersLockedOut(string iasoname, TimeSpan t, bool getusers)
        {
            List<ADUser> users = new List<ADUser>();
            List<OU> roots = DBHelper.GetSearchRoots(iasoname);
            foreach (OU ou in roots)
            {
                string path = "DC=kor,DC=ds,DC=cmil,DC=mil";
                if (!string.IsNullOrEmpty(ou.Path)) path = ou.Path.Substring(ou.Path.LastIndexOf('/') + 1);
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "kor.ds.cmil.mil", path, ContextOptions.SimpleBind, ADHelper.GetServiceAccount(), ADHelper.GetServicePassword());
                PrincipalSearcher ps = new PrincipalSearcher();
                UserPrincipal usr = new UserPrincipal(ctx);
                usr.Enabled = true;
                usr.AdvancedSearchFilter.AccountLockoutTime(DateTime.Now.Subtract(t), MatchType.GreaterThan);
                ps.QueryFilter = usr;
                (ps.GetUnderlyingSearcher() as DirectorySearcher).PageSize = 1000;
                PrincipalSearchResult<Principal> results = ps.FindAll();
                foreach (UserPrincipal u in results)
                {
                    if (getusers)
                    {
                        users.Add(ADHelper.GetUser(u.SamAccountName));
                    }
                    else
                    {
                        ADUser user = new ADUser();
                        user.Username = u.SamAccountName;
                        users.Add(user);
                    }
                }
            }
            return users;
        }


    }

	public class ADACEObjectTypes
	{
		public const string USER_LOGON_INFORMATION = "{5f202010-79a5-11d0-9020-00c04fc2d4cf}";
		public const string USER_ACCOUNT_RESTRICTIONS = "{4c164200-20c0-11d0-a768-00aa006e0529}";
		public const string SELF_MEMBERSHIP = "{C7407360-20BF-11D0-A768-00AA006E0529}";
		public const string VALIDATED_SPN = "{f3a64788-5306-11d1-a9c5-0000f80367c1}";
		public const string VALIDATED_DNS_HOST_NAME = "{72e39547-7b18-11d1-adef-00c04fd8d5cd}";
		public const string RESET_PASSWORD_GUID = "{00299570-246D-11D0-A768-00AA006E0529}";
		public const string DESCRIPTION_GUID = "{BF967950-0DE6-11D0-A285-00AA003049E2}";
		public const string SAMACCOUNTNAME_GUID = "{3E0ABFD0-126A-11D0-A060-00AA006C33ED}";
		public const string COMMONNAME_GUID = "{BF96793F-0DE6-11D0-A285-00AA003049E2}";
		public const string DISPLAYNAME_GUID = "{BF967953-0DE6-11D0-A285-00AA003049E2}";
        public const string ADS_OBJECT_WRITE_MEMBERS = "{BF9679C0-0DE6-11D0-A285-00AA003049E2}";
		public const string USER_CHANGE_PASSWORD = "{AB721A53-1E2f-11D0-9819-00AA0040529b}";
		public const string ALLOWED_TO_AUTHENTICATE = "{68B1D179-0D15-4d4f-AB71-46152E79A7BC}";
		public const string RECEIVE_AS = "{AB721A56-1E2f-11D0-9819-00AA0040529B}";
		public const string SEND_AS = "{AB721A54-1E2f-11D0-9819-00AA0040529B}";
		public const string COMPUTER = "{bf967a86-0de6-11d0-a285-00aa003049e2}";
		public const string UNKNOWN = "{00000000-0000-0000-0000-000000000000}";
	}


}