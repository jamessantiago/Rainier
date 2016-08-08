using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices;
using ActiveDs;

namespace Rainier
{
	/// <summary>
	/// A security group
	/// </summary>
	public class Group : Entry
	{
		#region Constructor
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="DirectoryEntry"></param>
		public Group(DirectoryEntry DirectoryEntry, PrincipalContext Context)
			: base(DirectoryEntry, Context)
		{
			this.DirectoryEntry = DirectoryEntry;
			this._Context = Context;
		}

		#endregion Constructor

		#region Properties

		private PrincipalContext _Context;

		/// <summary>
		/// distro property for this entry
		/// </summary>
		public virtual bool IsDistro
		{
			get
			{
				int GroupType = (int)GetValue("groupType");
				return (GroupType == (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_GLOBAL_GROUP || GroupType == (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_UNIVERSAL_GROUP);
			}
			set
			{
				if (value)
					SetValue("groupType", (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_GLOBAL_GROUP);
				else
					SetValue("groupType", (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_GLOBAL_GROUP | (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_SECURITY_ENABLED);
			}
		}

		/// <summary>
		/// universal group type property for this entry
		/// </summary>
		public virtual bool IsUniversal
		{
			get
			{
				int GroupType = (int)GetValue("groupType");
				return Convert.ToBoolean(GroupType & (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_UNIVERSAL_GROUP);
			}
			set
			{
				if (IsDistro)
				{
					if (value)
						SetValue("groupType", (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_UNIVERSAL_GROUP);
					else 
						SetValue("groupType", (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_GLOBAL_GROUP);
				}
				else
				{
					if (value)
						SetValue("groupType", (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_UNIVERSAL_GROUP | (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_SECURITY_ENABLED);
					else
						SetValue("groupType", (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_GLOBAL_GROUP | (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_SECURITY_ENABLED);
				}

			}
		}

		/// <summary>
		/// manager DN property for this entry
		/// </summary>
		public virtual string ManagerDN
		{
			get { return (string)GetValue("managedBy"); }
			set { SetValue("managedBy", value); }
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
		/// mail nick name property for this entry (example: first.last.us)
		/// </summary>
		public virtual string MailNickName
		{
			get { return (string)GetValue("mailNickName"); }
			set { SetValue("mailNickName", value); }
		}

		/// <summary>
		/// note property for this entry (example: first.last.us)
		/// </summary>
		public virtual string Notes
		{
			get { return (string)GetValue("info"); }
			set { SetValue("info", value); }
		}

		/// <summary>
		/// displayname property for this entry
		/// </summary>
		public virtual string DisplayName
		{
			get { return (string)GetValue("displayName"); }
			set { SetValue("displayName", value); }
		}

		#endregion Properties

		#region Public Methods

		#region AddUser

		/// <summary>
		/// Add user to this group
		/// </summary>
		/// <param name="user"></param>
		public void AddUser(string UserName)
		{
			using (var group = GetGroupPrinciple())
			{
				group.Members.Add(group.Context, IdentityType.SamAccountName, UserName);
				group.Save();
			}
		}

		#endregion AddUser

		#region RemoveUser

		/// <summary>
		/// Remove user from this group
		/// </summary>
		/// <param name="user"></param>
		public void RemoveUser(string UserName)
		{
			using (var group = GetGroupPrinciple())
			{
				group.Members.Remove(group.Context, IdentityType.SamAccountName, UserName);
				group.Save();
			}
		}

		#endregion RemoveUser

		#region CanUpdateMembers

		/// <summary>
		/// Determines if a user can update the members of this group
		/// </summary>
		/// <param name="UserSamAccountName"></param>
		/// <returns></returns>
		public bool CanUpdateMembers(string UserSamAccountName)
		{
			foreach (IADsAccessControlEntry entry in AccessControlList)
			{
				if (entry.Trustee.Contains(UserSamAccountName) && entry.ObjectType == Tools.ADACEObjectTypes.ADS_OBJECT_WRITE_MEMBERS)
					return true;
			}
			return false;
		}

		#endregion CanUpdateMembers

		#region SetUpdateMembersPriviledge

		/// <summary>
		/// Sets the ace to write to members for this group
		/// </summary>
		/// <param name="UserName"></param>
		public void SetUpdateMembersPriviledge(string UserName)
		{
			ActiveDirectoryAccessRule rule = new ActiveDirectoryAccessRule(new System.Security.Principal.NTAccount("kor" + "\\" + UserName), 
				ActiveDirectoryRights.WriteProperty, 
				System.Security.AccessControl.AccessControlType.Allow, 
				new Guid(Tools.ADACEObjectTypes.ADS_OBJECT_WRITE_MEMBERS));
			DirectoryEntry.ObjectSecurity.AddAccessRule(rule);
		}

		#endregion SetUpdateMembersPriviledge

		#region HasMember

		/// <summary>
		/// Determines if a user is a member of a group
		/// </summary>
		/// <param name="UserName"></param>
		/// <returns></returns>
		public bool HasMember(string UserName)
		{
			using (var group = GetGroupPrinciple())
			{
				return group.Members.Any(d => d.SamAccountName == UserName);
			}
		}

		#endregion HasMember

		#region Delete

		/// <summary>
		/// Delete this entry
		/// </summary>
		public void Delete()
		{
			using (var group = GetGroupPrinciple())
			{
				group.Delete();
			}
		}

		#endregion Delete

		#region Get Group Principle
		/// <summary>
		/// Gets the user principle which can be used to perform advanced functions and access properties
		/// </summary>
		public GroupPrincipal GetGroupPrinciple()
		{
			int tryCount = 0;
			GroupPrincipal group = null; ;
			while (tryCount < 10)
			{
				group = GroupPrincipal.FindByIdentity(_Context, IdentityType.DistinguishedName, this.DistinguishedName);
				if (group != null)
					break;
				System.Threading.Thread.Sleep(1000);
				tryCount++;
			}
			if (group == null)
				throw new Exception("Failed to retrieve GroupPrincipal for " + this.DistinguishedName);
			return group;
		}
		#endregion Get User Principle

		#endregion Public Methods
	}
}
