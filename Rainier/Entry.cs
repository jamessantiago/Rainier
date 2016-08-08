#region Usings
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using ActiveDs;
#endregion

namespace Rainier
{

	/// <summary>
	/// Directory entry class
	/// </summary>
	public class Entry : IDisposable
	{
		#region Constructors
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="DirectoryEntry">Directory entry for the item</param>
		public Entry(DirectoryEntry DirectoryEntry, PrincipalContext Context)
		{
			this.DirectoryEntry = DirectoryEntry;
			this._Context = Context;
		}
		#endregion

		#region Properties

		private PrincipalContext _Context;

		public PrincipalContext Context { get { return _Context; } }

		/// <summary>
		/// Actual base directory entry
		/// </summary>
		public virtual DirectoryEntry DirectoryEntry { get; set; }

		/// <summary>
		/// distinguished name property for this entry
		/// </summary>
		public virtual string DistinguishedName
		{
			get { return (string)GetValue("distinguishedName"); }
			set { SetValue("distinguishedName", value); }
		}

		/// <summary>
		/// MemberOf property for this entry
		/// </summary>
		public virtual List<string> MemberOf
		{
			get
			{
				List<string> Values = new List<string>();
				PropertyValueCollection Collection = DirectoryEntry.Properties["memberof"];
				foreach (object Item in Collection)
				{
					Values.Add((string)Item);
				}
				return Values;
			}
		}

		/// <summary>
		/// samaccountname property for this entry
		/// </summary>
		public virtual string SamAccountName
		{
			get { return (string)GetValue("sAMAccountName"); }
			set { SetValue("sAMAccountName", value); }
		}

		/// <summary>
		/// cn property for this entry
		/// </summary>
		public virtual string CN
		{
			get { return (string)GetValue("cn"); }
			set { SetValue("cn", value); }
		}

		/// <summary>
		/// name property for this entry
		/// </summary>
		public virtual string Name
		{
			get { return (string)GetValue("name"); }
			set { SetValue("name", value); }
		}

		/// <summary>
		/// description property for this entry
		/// </summary>
		public virtual string Description
		{
			get { return (string)GetValue("description"); }
			set { SetValue("description", value); }
		}

		/// <summary>
		/// distro property for this entry
		/// </summary>
		public virtual bool IsEntryDistro
		{
			get
			{
				int GroupType = (int)GetValue("groupType");
				return (GroupType == (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_GLOBAL_GROUP || GroupType == (int)ADS_GROUP_TYPE_ENUM.ADS_GROUP_TYPE_UNIVERSAL_GROUP);
			}
		}

		/// <summary>
		/// the object category property for this entry
		/// </summary>
		public virtual string ObjectCategory
		{
			get { return (string)GetValue("objectCategory"); }
			set { SetValue("objectCategory", value); }
		}

		/// <summary>
		/// disable property for this entry
		/// </summary>
		public virtual bool IsEntryDisabled
		{
			get
			{
				int uacVal = (int)GetValue("userAccountControl");
				return Convert.ToBoolean(uacVal & (int)ADS_USER_FLAG.ADS_UF_ACCOUNTDISABLE);
			}
			set
			{
				int uacVal = (int)GetValue("userAccountControl");
				if (value)
					SetValue("userAccountControl", (uacVal | (int)ADS_USER_FLAG.ADS_UF_ACCOUNTDISABLE));
				else
					SetValue("userAccountControl", (uacVal | ~(int)ADS_USER_FLAG.ADS_UF_ACCOUNTDISABLE));
			}
		}

		/// <summary>
		/// integer value of useraccountcontrol, it is not recommended to set
		/// the value directly as you may be removing flags set in the process
		/// </summary>
		public virtual int RawUserAccountControl
		{
			get
			{
				return (int)GetValue("userAccountControl");
			}
			set
			{
					SetValue("userAccountControl", value);
			}
		}

		/// <summary>
		/// nt security descriptor for this entry
		/// </summary>
		public virtual IADsSecurityDescriptor SecurityDescriptor
		{
			get { return (IADsSecurityDescriptor)GetValue("ntSecurityDescriptor"); }
			set { SetValue("ntSecurityDescriptor", value); }
		}

		/// <summary>
		/// the discretionary acl from this entries security descriptor
		/// </summary>
		public virtual IADsAccessControlList AccessControlList
		{
			get { return (IADsAccessControlList)SecurityDescriptor.DiscretionaryAcl; }
		}

		/// <summary>
		/// Get's the distinguished name of this entry's parent
		/// </summary>
		public virtual string ParentDN
		{
			get { return Rainier.Tools.GetDNFromPath(this.DirectoryEntry.Parent.Path); }
		}

		/// <summary>
		/// created date property for this entry
		/// </summary>
		public virtual DateTime ObjectCreated
		{
			get
			{
				try
				{
					var results = GetValue("whenCreated") != null ? (DateTime)GetValue("whenCreated") : DateTime.MinValue;
					return results;
				}
				catch
				{
					return DateTime.MinValue;
				}
			}
		}

		/// <summary>
		/// last changed date property for this entry
		/// </summary>
		public virtual DateTime ObjectChanged
		{
			get
			{
				try
				{
					var results = GetValue("whenChanged") != null ? (DateTime)GetValue("whenChanged") : DateTime.MinValue;
					return results;
				}
				catch
				{
					return DateTime.MinValue;
				}
			}
		}

		#endregion

		#region Public Functions
		/// <summary>
		/// Saves any changes that have been made
		/// </summary>
		public virtual void Save()
		{
			if (DirectoryEntry == null)
				throw new NullReferenceException("DirectoryEntry shouldn't be null");
			DirectoryEntry.CommitChanges();
		}

		/// <summary>
		/// Gets a value from the entry
		/// </summary>
		/// <param name="Property">Property you want the information about</param>
		/// <returns>an object containing the property's information</returns>
		public virtual object GetValue(string Property)
		{
			PropertyValueCollection Collection = DirectoryEntry.Properties[Property];
			return Collection != null ? Collection.Value : null;
		}

		/// <summary>
		/// Gets a value from the entry
		/// </summary>
		/// <param name="Property">Property you want the information about</param>
		/// <param name="Index">Index of the property to return</param>
		/// <returns>an object containing the property's information</returns>
		public virtual object GetValue(string Property, int Index)
		{
			PropertyValueCollection Collection = DirectoryEntry.Properties[Property];
			return Collection != null ? Collection[Index] : null;
		}

		/// <summary>
		/// Sets a property of the entry to a specific value
		/// </summary>
		/// <param name="Property">Property of the entry to set</param>
		/// <param name="Value">Value to set the property to</param>
		public virtual void SetValue(string Property, object Value)
		{
			PropertyValueCollection Collection = DirectoryEntry.Properties[Property];
			if (Collection != null)
				Collection.Value = Value;
		}

		/// <summary>
		/// Sets a property of the entry to a specific value
		/// </summary>
		/// <param name="Property">Property of the entry to set</param>
		/// <param name="Index">Index of the property to set</param>
		/// <param name="Value">Value to set the property to</param>
		public virtual void SetValue(string Property, int Index, object Value)
		{
			PropertyValueCollection Collection = DirectoryEntry.Properties[Property];
			if (Collection != null)
				Collection[Index] = Value;
		}

		/// <summary>
		/// Move entry to different directory
		/// </summary>
		/// <param name="Path">the LDAP path to move to</param>
		public virtual void MoveTo(Directory dir)
		{
			this.DirectoryEntry.MoveTo(dir.Entry);
		}

		#endregion

		#region IDisposable Members

		public void Dispose()
		{
			if (DirectoryEntry != null)
			{
				DirectoryEntry.Dispose();
				DirectoryEntry = null;
			}
		}

		#endregion
	}
}