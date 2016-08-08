using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security;
using Rainier;
using ActiveDs;
using TSUSEREXLib;

namespace Rainier
{
	public class User : Entry
	{
		#region Constructor
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="DirectoryEntry">Directory entry for the user</param>
		public User(DirectoryEntry DirectoryEntry, PrincipalContext Context)
			: base(DirectoryEntry, Context)
		{
			this.DirectoryEntry = DirectoryEntry;
			this._Context = Context;
		}

		/// <summary>
		/// Creates a user class
		/// </summary>
		/// <param name="entry">An entry that is a user object</param>
		public User(Entry entry)
			: base(entry.DirectoryEntry, entry.Context)
		{
			if (!entry.ObjectCategory.Contains("Person"))
				throw new Exception("Provided entry is not a user object");
			this.DirectoryEntry = entry.DirectoryEntry;
			this._Context = entry.Context;
		}
		#endregion Constructor

		#region Properties

		private PrincipalContext _Context;

		#region Name
		/// <summary>
		/// givenname property for this entry (First Name)
		/// </summary>
		public virtual string GivenName
		{
			get { return (string)GetValue("givenname"); }
			set { SetValue("givenname", value); }
		}

		/// <summary>
		/// initials property for this entry
		/// </summary>
		public virtual string Initials
		{
			get { return (string)GetValue("initials"); }
			set { SetValue("initials", value); }
		}

		/// <summary>
		/// surname property for this entry (Last Name)
		/// </summary>
		public virtual string Surname
		{
			get { return (string)GetValue("sn"); }
			set { SetValue("sn", value); }
		}
		#endregion Name

		#region General

		/// <summary>
		/// display name property for this entry
		/// </summary>
		public virtual string DisplayName
		{
			get { return (string)GetValue("displayname"); }
			set { SetValue("displayname", value); }
		}

		//inherits description

		/// <summary>
		/// office property for this entry
		/// </summary>
		public virtual string Office
		{
			get { return (string)GetValue("physicaldeliveryofficename"); }
			set { SetValue("physicaldeliveryofficename", value); }
		}

		/// <summary>
		/// telephone number property for this entry
		/// </summary>
		public virtual string TelephoneNumber
		{
			get { return (string)GetValue("telephonenumber"); }
			set { SetValue("telephonenumber", value); }
		}

		/// <summary>
		/// Email property for this entry
		/// </summary>
		public virtual string Email
		{
			get { return (string)GetValue("mail"); }
			set { SetValue("mail", value); }
		}

		/// <summary>
		/// webpage property for this entry
		/// </summary>
		public virtual string WebPage
		{
			get { return (string)GetValue("wWWHomePage"); }
			set { SetValue("wWWHomePage", value); }
		}

		public virtual byte[] Photo
		{
			get { return (byte[])GetValue("thumbnailphoto"); }
			set { SetValue("thumbnailphoto", value); }
		}

		#endregion General

		#region Address

		/// <summary>
		/// street address property for this entry
		/// </summary>
		public virtual string StreetAddress
		{
			get { return (string)GetValue("streetAddress"); }
			set { SetValue("streetAddress", value); }
		}

		/// <summary>
		/// city property for this entry
		/// </summary>
		public virtual string City
		{
			get { return (string)GetValue("l"); }
			set { SetValue("l", value); }
		}

		/// <summary>
		/// state property for this entry
		/// </summary>
		public virtual string State
		{
			get { return (string)GetValue("st"); }
			set { SetValue("st", value); }
		}

		/// <summary>
		/// zip property for this entry
		/// </summary>
		public virtual string Zip
		{
			get { return (string)GetValue("postalCode"); }
			set { SetValue("postalCode", value); }
		}

		/// <summary>
		/// country code property for this entry
		/// </summary>
		public virtual int CountryCode
		{
			get { return (int)GetValue("countrycode"); }
			set { SetValue("countrycode", value); }
		}

		#endregion Address

		#region Account

		//samaccountname is inherited

		/// <summary>
		/// userPrincipalName property for this entry
		/// </summary>
		public virtual string UserPrincipalName
		{
			get { return (string)GetValue("userPrincipalName"); }
			set { SetValue("userPrincipalName", value); }
		}

		/// <summary>
		/// account expiration date property for this entry
		/// </summary>
		public virtual DateTime ExpirationDate
		{
			get
			{
				try
				{
					var result = GetValue("AccountExpires") != null ? Tools.ToDateTime((LargeInteger)(GetValue("AccountExpires"))) : DateTime.MinValue;
					return result;
				}
				catch
				{
					return DateTime.MinValue;
				}
			}
			set { SetValue("AccountExpires", Tools.ToLargeInteger(value)); }
		}

		/// <summary>
		/// enabled status for this user
		/// </summary>
		public virtual bool Enabled
		{
			get
			{
				using (var user = GetUserPrinciple())
				{
					return user.Enabled ?? false;
				}
			}
			set
			{
				using (var user = GetUserPrinciple())
				{
					user.Enabled = value;
					user.Save();
				}
			}
		}

		//IsDisabled is inherited

		/// <summary>
		/// last logon timestamp date property for this entry
		/// </summary>
		public virtual DateTime LastLogonTimestamp
		{
			get
			{
				try
				{
					var results = GetValue("lastLogonTimestamp") != null ? Tools.ToDateTime((LargeInteger)(GetValue("lastLogonTimestamp"))) : DateTime.MinValue;
					return results;
				}
				catch
				{
					return DateTime.MinValue;
				}
			}
			set { SetValue("lastLogonTimestamp", Tools.ToLargeInteger(value)); }
		}

		/// <summary>
		/// last lockout timestamp date property for this entry
		/// </summary>
		public virtual DateTime LockoutTimestamp
		{
			get
			{
				try
				{
					var results = GetValue("lockoutTime") != null ? Tools.ToDateTime((LargeInteger)(GetValue("lastLogonTimestamp"))) : DateTime.MinValue;
					return results;
				}
				catch
				{
					return DateTime.MinValue;
				}
			}
		}

		/// <summary>
		/// password last set property for this entry
		/// </summary>
		public virtual DateTime PasswordLastSet
		{
			get
			{
				try
				{
					var results = GetValue("pwdLastSet") != null ? Tools.ToDateTime((LargeInteger)(GetValue("pwdLastSet"))) : DateTime.MinValue;
					return results;
				}
				catch
				{
					return DateTime.MinValue;
				}
			}
		}

		#endregion Account

		#region Profile

		/// <summary>
		/// logon script property for this entry
		/// </summary>
		public virtual string LogonScript
		{
			get { return (string)GetValue("scriptPath"); }
			set { SetValue("scriptPath", value); }
		}

		/// <summary>
		/// home drive property for this entry
		/// </summary>
		public virtual string HomeDrive
		{
			get { return (string)GetValue("homeDrive"); }
			set { SetValue("homeDrive", value); }
		}

		/// <summary>
		/// home directory property for this entry
		/// </summary>
		public virtual string HomeDirectory
		{
			get { return (string)GetValue("homeDirectory"); }
			set { SetValue("homeDirectory", value); }
		}

		#endregion Profile

		#region Telephones

		/// <summary>
		/// ip phone property for this entry
		/// </summary>
		public virtual string IpPhone
		{
			get { return (string)GetValue("ipPhone"); }
			set { SetValue("ipPhone", value); }
		}

		/// <summary>
		/// mobile phone property for this entry
		/// </summary>
		public virtual string MobilePhone
		{
			get { return (string)GetValue("mobile"); }
			set { SetValue("mobile", value); }
		}

		#endregion Telephones

		#region Organization

		/// <summary>
		/// rank property for this entry
		/// </summary>
		public virtual string Rank
		{
			get { return (string)GetValue("title"); }
			set { SetValue("title", value); }
		}

		/// <summary>
		/// department property for this entry
		/// </summary>
		public virtual string Department
		{
			get { return (string)GetValue("department"); }
			set { SetValue("department", value); }
		}

		/// <summary>
		/// company property for this entry
		/// </summary>
		public virtual string Company
		{
			get { return (string)GetValue("company"); }
			set { SetValue("company", value); }
		}

		#endregion Organization

		#region Exchange

		/// <summary>
		/// is hidden from gal property for this entry
		/// </summary>
		public virtual bool IsHiddenFromGal
		{
			get { return Convert.ToBoolean(GetValue("msExchHideFromAddressLists") ?? false); }
			set { SetValue("msExchHideFromAddressLists", value); }
		}

		/// <summary>
		/// mail nick name property for this entry (example: first.last.us)
		/// </summary>
		public virtual string MailNickName
		{
			get { return (string)GetValue("mailNickName"); }
			set { SetValue("mailNickName", value); }
		}

		//email is under general

		/// <summary>
		/// exchange server property for this entry (msExchHomeServerName)
		/// </summary>
		public virtual string ExchangeServer
		{
			get { return (string)GetValue("msExchHomeServerName"); }
			set { SetValue("msExchHomeServerName", value); }
		}

		/// <summary>
		/// exchange server mta property for this entry (homeMTA)
		/// </summary>
		public virtual string ExchangeServerMta
		{
			get { return (string)GetValue("homeMTA"); }
			set { SetValue("homeMTA", value); }
		}

		/// <summary>
		/// exchange server mdb property for this entry (homeMTA)
		/// </summary>
		public virtual string ExchangeServerMdb
		{
			get { return (string)GetValue("homeMDB"); }
			set { SetValue("homeMDB", value); }
		}

		/// <summary>
		/// legacy exchange dn property for this entry (legacyExchangeDN)
		/// </summary>
		public virtual string LegacyExchangeDn
		{
			get { return (string)GetValue("legacyExchangeDN"); }
			set { SetValue("legacyExchangeDN", value); }
		}

		/// <summary>
		/// Retrieves the msExchMailboxGuid property for the entry
		/// </summary>
		public virtual string ExchangeMailboxGuid
		{
			get { 
					var guid = (Byte[])GetValue("msExchMailboxGuid");
					Guid result = new Guid(guid);
					return result.ToString();
				}
			set {
				string data = value;

				Guid newguid = new Guid(data);

				//remove nonhex characters
				string guid = newguid.ToString().Replace("{", "").Replace("}", "").Replace("-", "");
				//reverse first three sequences
				string firstseq = guid.Substring(4, 2);
				string secondseq = guid.Substring(2, 2);
				string thirdseq = guid.Substring(0, 2);
				guid = firstseq + secondseq + thirdseq + guid.Remove(0, 6);
				//add backslashes every two chars
				string finalguid = "";
				for (int i = 0; i < guid.Length; i += 2)
					finalguid += guid.Substring(i, 2) + "\\";
				finalguid = finalguid.TrimEnd('\\');

				SetValue("msExchMailboxGuid", finalguid); }
		}

		/// <summary>
		/// Gets the native object of this user
		/// </summary>
		public virtual object NativeObject
		{
			get { return this.DirectoryEntry.NativeObject; }
		}

		#endregion Exchange

		#region Terminal Services

		/// <summary>
		/// Returns the terminal services home directory attribute from the userParameters blob
		/// </summary>
		public virtual string TerminalServicesHomeDirectory
		{
			get { try { return GetTSUser(this).TerminalServicesHomeDirectory; } catch { return ""; } }
			set { GetTSUser(this).TerminalServicesHomeDirectory = value; }
		}

		/// <summary>
		/// Returns the terminal services profile path attribute from the userParameters blob
		/// </summary>
		public virtual string TerminalServicesProfilePath
		{
			get { try { return GetTSUser(this).TerminalServicesProfilePath; } catch { return ""; } }
			set { GetTSUser(this).TerminalServicesProfilePath = value; }
		}

		/// <summary>
		/// Returns the terminal services home drive attribute from the userParameters blob
		/// </summary>
		public virtual string TerminalServicesHomeDrive
		{
			get { try { return GetTSUser(this).TerminalServicesHomeDrive; } catch { return ""; } }
			set { GetTSUser(this).TerminalServicesHomeDrive = value; }
		}

		#endregion Terminal Services

		#region Extended Properties

		/// <summary>
		/// unclass email property for this entry
		/// </summary>
		public virtual string UnclassEmail
		{
			get { return (string)GetValue("extensionAttribute5"); }
			set { SetValue("extensionAttribute5", value); }
		}

		/// <summary>
		/// branch id property for this entry
		/// </summary>
		public virtual string BranchID
		{
			get { return (string)GetValue("extensionAttribute7"); }
			set { SetValue("extensionAttribute7", value); }
		}

		/// <summary>
		/// nationality property for this entry
		/// </summary>
		public virtual string Nationality
		{
			get { return (string)GetValue("extensionAttribute8"); }
			set { SetValue("extensionAttribute8", value); }
		}

		/// <summary>
		/// nationality property for this entry
		/// </summary>
		public virtual DateTime CreatedDate
		{
			get
			{
				DateTime result;
				DateTime.TryParse((string)GetValue("extensionAttribute10"), out result);
				return result;
			}
			set { SetValue("extensionAttribute10", value.ToString()); }
		}

		/// <summary>
		/// clearance property for this entry
		/// </summary>
		public virtual string Clearance
		{
			get { return (string)GetValue("extensionAttribute11"); }
			set { SetValue("extensionAttribute11", value); }
		}

		/// <summary>
		/// clearance date property for this entry
		/// </summary>
		public virtual DateTime ClearanceDate
		{
			get {
				DateTime result;
				DateTime.TryParse((string)GetValue("extensionAttribute12"), out result);
				return result; }
			set { SetValue("extensionAttribute12", value.ToString()); }
		}

		/// <summary>
		/// Title property for this entry
		/// </summary>
		public virtual string Title
		{
			get { return (string)GetValue("extensionAttribute13"); }
			set { SetValue("extensionAttribute13", value); }
		}

		/// <summary>
		/// msc (major subordinate command) property for this entry
		/// </summary>
		public virtual string MSC
		{
			get { return (string)GetValue("extensionAttribute9"); }
			set { SetValue("extensionAttribute9", value); }
		}

        /// <summary>
		/// userWorkstations property for this entry
		/// </summary>
        public virtual List<string> LogOnToWorkstations
        {
            get {
                var workstations = (string)GetValue("userWorkstations");
                return string.IsNullOrEmpty(workstations) ? new List<string>() : workstations.Split(',').ToList();
            }
            set {
                SetValue("userWorkstations", string.Join(",", value));
            }
        }

		#endregion Extended Properties

		

		#endregion Properties

		#region Public Methods

		#region Change Password
		/// <summary>
		/// Changes account password
		/// </summary>
		/// <param name="oldPassword">The old password</param>
		/// <param name="newPassword">The new password</param>
		public void ChangePassword(string oldPassword, string newPassword)
		{
			using (var user = GetUserPrinciple())
			{
				user.ChangePassword(oldPassword, newPassword);
			}
		}
		#endregion Change Password

		#region Set Pasword
		/// <summary>
		/// Sets a new password for the user
		/// </summary>
		/// <param name="newPassword">The new password</param>
		public void SetPassword(string newPassword)
		{
			using (var user = GetUserPrinciple())
			{
				user.SetPassword(newPassword);
			}
		}
		#endregion Set Pasword

		#region Expire Pasword
		/// <summary>
		/// Expires password for the user
		/// </summary>
		public void ExpirePassword()
		{
			using (var user = GetUserPrinciple())
			{
				user.ExpirePasswordNow();
			}
		}
		#endregion Set Pasword

		#region Unlock
		/// <summary>
		///Unlocks user account
		/// </summary>
		public void Unlock()
		{
			using (var user = GetUserPrinciple())
			{
				user.UnlockAccount();		
			}
		}
		#endregion Unlock

		#region IsLocked
		public bool IsLocked()
		{
			using (var user = GetUserPrinciple())
			{
				return user.IsAccountLockedOut();
			}
		}
		#endregion IsLocked

		#region IsMemberOf

		/// <summary>
		/// Determines if this user is a member of a group
		/// </summary>
		/// <param name="GroupName"></param>
		/// <returns></returns>
		public bool IsMemberof(string GroupName)
		{
			foreach (string group in MemberOf)
			{
				if (group.Contains(GroupName))
					return true;
			}
			return false;
		}

		#endregion IsMemberOf

		#region AddGroup
		/// <summary>
		/// Add a user to a security group
		/// </summary>
		/// <param name="UserName"></param>
		/// <param name="GroupName"></param>
		public void AddGroup(Group group)
		{
			group.AddUser(this.SamAccountName);
		}

		#endregion AddGroup

		#region RemoveGroup

		/// <summary>
		/// Remove a group from this user's members
		/// </summary>
		/// <param name="group"></param>
		public void RemoveGroup(Group group)
		{
			group.RemoveUser(this.SamAccountName);
		}

		#endregion RemoveGroup

		#region Delete
		/// <summary>
		/// Deletes the account
		/// </summary>
		public void Delete()
		{
			using (var user = GetUserPrinciple())
			{
				user.Delete();
			}
		}
		#endregion Delete

		#region Get User Principle
		/// <summary>
		/// Gets the user principle which can be used to perform advanced functions and access properties
		/// </summary>
		public UserPrincipal GetUserPrinciple()
		{
			
			int tryCount = 0;
			UserPrincipal user = null;
			while (tryCount < 10)
			{
				user = UserPrincipal.FindByIdentity(_Context, IdentityType.SamAccountName, this.SamAccountName);
				if (user != null)
					break;
				System.Threading.Thread.Sleep(1000);
				tryCount++;
			}
			if (user == null)
				throw new Exception("Failed to retrieve UserPrincipal for " + this.SamAccountName);
			return user;
		}
		#endregion Get User Principle

		#endregion Public Methods

		#region private Methods

		private IADsTSUserEx GetTSUser(User user)
		{
			return (IADsTSUserEx)user.NativeObject;
		}

		#endregion private Methods
	}
	
}
