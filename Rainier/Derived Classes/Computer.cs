using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using ActiveDs;

namespace Rainier
{

	/// <summary>
	/// A computer
	/// </summary>
	public class Computer : Entry
	{
		#region Constructor
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="DirectoryEntry"></param>
		public Computer(DirectoryEntry DirectoryEntry, PrincipalContext Context)
			: base(DirectoryEntry, Context)
		{
			this.DirectoryEntry = DirectoryEntry;
			this._Context = Context;
		}

		#endregion Constructor

		#region Properties

		/// <summary>
		/// enabled status for this computer
		/// </summary>
		public virtual bool Enabled
		{
			get
			{
				using (var computer = GetComputerPrinciple())
				{
					return computer.Enabled ?? false;
				}
			}
			set
			{
				using (var computer = GetComputerPrinciple())
				{
					computer.Enabled = value;
					computer.Save();
				}
			}
		}

		/// <summary>
		/// last logon timestamp date property for this entry
		/// </summary>
		public virtual DateTime LastLogonTimestamp
		{
			get { return Tools.ToDateTime((LargeInteger)GetValue("lastLogonTimestamp")); }
			set { SetValue("lastLogonTimestamp", Tools.ToLargeInteger(value)); }
		}

		/// <summary>
		/// info property for this entry
		/// </summary>
		public virtual string ExtendedInfo
		{
			get { return (string)GetValue("info"); }
			set { SetValue("info", value); }
		}

		/// <summary>
		/// operating system property
		/// </summary>
		public virtual string OperatingSystem
		{
			get { return (string)GetValue("operatingSystem"); }
			set { SetValue("operatingSystem", value); }
		}

		/// <summary>
		/// operating system service pack property
		/// </summary>
		public virtual string OperatingSystemServicePack
		{
			get { return (string)GetValue("operatingSystemServicePack"); }
			set { SetValue("operatingSystemServicePack", value); }
		}

		/// <summary>
		/// operating system version property
		/// </summary>
		public virtual string OperatingSystemVersion
		{
			get { return (string)GetValue("operatingSystemVersion"); }
			set { SetValue("operatingSystemVersion", value); }
		}

		#endregion Properties

		#region Public Methods

		private PrincipalContext _Context;

		#region SetInitialUserAccountControl

		/// <summary>
		/// Sets the intial useraccountcontrol for the computer (pass not required | workstation trust account)
		/// </summary>
		public void SetInitialUserAccountControl()
		{
			SetValue("userAccountControl", 4128);
		}

		#endregion SetInitialUserAccountControl

		#region SetJoinPermissions

		/// <summary>
		/// Sets the permission to join this computer to the domain to a trustee such as domain\user or Authenticated Users
		/// </summary>
		/// <param name="Trustee"></param>
		public void SetJoinPermissions(string Trustee)
		{
			//create a temporary acl
			IADsAccessControlList acl = AccessControlList;			
			
			//Gets aces from tools
			Tools.ADACEComputerJoinPermissions acllist = new Tools.ADACEComputerJoinPermissions(Trustee);
			foreach (IADsAccessControlEntry ace in acllist.ace_writeaccountrestrictions)
			{
				acl.AddAce(ace);
			}

			//Update the security descriptor with the new ACL
			IADsSecurityDescriptor sd = SecurityDescriptor;
			sd.DiscretionaryAcl = acl;
			SecurityDescriptor = sd;
		}

		#endregion SetJoinPermissions

		public void ReplacePermisions(Computer baseComputer)
		{
			//create a temporary acl
			IADsAccessControlList acl = AccessControlList;
			IADsAccessControlList baseacl = baseComputer.AccessControlList;
			
			IADsSecurityDescriptor sd = SecurityDescriptor;
			sd.DiscretionaryAcl = baseacl;
			SecurityDescriptor = sd;
		}

		#region Reset

		/// <summary>
		/// resets the computer account
		/// </summary>
		public void Reset(string password)
		{
			using (var comp = GetComputerPrinciple())
			{				
				comp.SetPassword(password);
				comp.Save();
			}
		}

		#endregion Reset

		#region Delete

		/// <summary>
		/// Deletes the computer account
		/// </summary>
		public void Delete()
		{
			GetComputerPrinciple().Delete();
		}

		#endregion Delete

		#region GetComputerPrinciple

		/// <summary>
		/// the computer principle
		/// </summary>
		public virtual ComputerPrincipal GetComputerPrinciple()
		{
			int tryCount = 0;
			ComputerPrincipal comp = null; ;
			while (comp == null && tryCount < 10)
			{
				comp = ComputerPrincipal.FindByIdentity(_Context, IdentityType.SamAccountName, this.SamAccountName);
				System.Threading.Thread.Sleep(1000);
				tryCount++;
			}
			if (comp == null)
				throw new Exception("Failed to retrieve ComputerPrincipal for " + this.SamAccountName);
			return comp;
		}

		#endregion GetComputerPrinciple

		#endregion Public Methods
	}
}
