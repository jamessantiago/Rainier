using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace Rainier
{
	/// <summary>
	/// A distro group
	/// </summary>
	public class DistroGroup : Entry
	{
		#region Constructor
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="DirectoryEntry"></param>
		public DistroGroup(DirectoryEntry DirectoryEntry, PrincipalContext Context)
			: base(DirectoryEntry, Context)
		{
			this.DirectoryEntry = DirectoryEntry;
			this._Context = Context;
		}

		#endregion Constructor

		#region Properties

		private PrincipalContext _Context;

		#endregion Properties

		#region Public Methods

		#region AddUser

		/// <summary>
		/// Add user to this group
		/// </summary>
		/// <param name="user"></param>
		public void AddUser(User user)
		{
			using (var group = GetGroupPrinciple())
			{
				group.Members.Add(user.GetUserPrinciple());
			}
		}

		#endregion AddUser

		#region RemoveUser

		/// <summary>
		/// Remove user from this group
		/// </summary>
		/// <param name="user"></param>
		public void RemoveUser(User user)
		{
			using (var group = GetGroupPrinciple())
			{
				group.Members.Remove(user.GetUserPrinciple());
			}
		}

		#endregion RemoveUser

		#region Get Group Principle
		/// <summary>
		/// Gets the user principle which can be used to perform advanced functions and access properties
		/// </summary>
		public GroupPrincipal GetGroupPrinciple()
		{
			using (var context = new PrincipalContext(ContextType.Domain))
			{
				return GroupPrincipal.FindByIdentity(context, IdentityType.DistinguishedName, this.DistinguishedName);
			}
		}
		#endregion Get User Principle

		#endregion Public Methods
	}

}
