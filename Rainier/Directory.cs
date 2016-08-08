using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.DirectoryServices.AccountManagement;
using ActiveDs;
using Rainier.Generics;
using System.DirectoryServices.ActiveDirectory;

namespace Rainier
{
	/// <summary>
	/// Class for helping with AD
	/// </summary>
	public class Directory : IDisposable
	{
		#region Constructors
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="UserName">User name used to log in</param>
		/// <param name="Password">Password used to log in</param>
		/// <param name="Path">Path of the LDAP server</param>
		/// <param name="Query">Query to use in the search</param>
		/// <param name="PropertiesToLoad">Specify exact entry properties to load for searches.  Ensure this is used to speed up queries otherwise all properties will be loaded and slow down results.</param>
		public Directory(string Query, string UserName, string Password, string Path, params string[] PropertiesToLoad)
		{
			_Entry = new DirectoryEntry(Path, UserName, Password, AuthenticationTypes.Secure);			
			this.Path = Path;
			this.UserName = UserName;
			this.Password = Password;
			this.Query = Query;
			this.PropertiesToLoad = PropertiesToLoad;			
			Searcher = new DirectorySearcher(_Entry);
			Searcher.Filter = Query;
			Searcher.PageSize = 1000;
			Searcher.PropertiesToLoad.AddRange(PropertiesToLoad);
			string[] serverinfo = Tools.ReverseRootLDAPPath(this.Path);
			_Domain = serverinfo[0];
			_DomainDN = serverinfo[1];
			this._Context = new PrincipalContext(ContextType.Domain, serverinfo[0], serverinfo[1], this.UserName, this.Password);
		}
		#endregion

		#region Public Functions

		#region Authenticate

		/// <summary>
		/// Checks to see if the person was authenticated
		/// </summary>
		/// <returns>true if they were authenticated properly, false otherwise</returns>
		public virtual bool Authenticate()
		{
			try
			{
				if (!_Entry.Guid.ToString().ToLower().Trim().Equals(""))
					return true;
			}
			catch { }
			return false;
		}
		#endregion

		#region Close

		/// <summary>
		/// Closes the directory
		/// </summary>
		public virtual void Close()
		{
			Entry.Close();
		}

		#endregion

		#region Entry Functions

		#region FindAll

		/// <summary>
		/// Finds all entries that match the query
		/// </summary>
		/// <returns>A list of all entries that match the query</returns>
		public virtual List<Entry> FindAll()
		{
			List<Entry> ReturnedResults = new List<Entry>();
			using (SearchResultCollection Results = Searcher.FindAll())
			{
				foreach (SearchResult Result in Results)
					ReturnedResults.Add(new Entry(Result.GetDirectoryEntry(), Context));
			}
			return ReturnedResults;
		}

        /// <summary>
		/// Finds all entries that match the query
		/// </summary>
		/// <returns>A list of all entries that match the query</returns>
		public virtual List<Entry> FindAll(string Filter)
        {
            Searcher.Filter = Filter;
            List<Entry> ReturnedResults = new List<Entry>();
            using (SearchResultCollection Results = Searcher.FindAll())
            {
                foreach (SearchResult Result in Results)
                    ReturnedResults.Add(new Entry(Result.GetDirectoryEntry(), Context));
            }
            return ReturnedResults;
        }

        #endregion

        #region FindOne

        /// <summary>
        /// Finds one entry that matches the query
        /// </summary>
        /// <returns>A single entry matching the query</returns>
        public virtual Entry FindOne()
		{
			return new Entry(Searcher.FindOne().GetDirectoryEntry(), Context);
		}

		#endregion

		#region FindUsersAndGroups

		/// <summary>
		/// Finds all users and groups
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all users and groups meeting the specified Filter</returns>
		public virtual List<Entry> FindUsersAndGroups(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&(|(&(objectClass=Group)(objectCategory=Group))(&(objectClass=User)(objectCategory=Person)))({0}))", Filter);
			Searcher.Filter = Filter;
			return FindAll();
		}

		#endregion

		#region FindUsersAndGroupsAndComputers

		/// <summary>
		/// Finds all users and groups
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all users and groups meeting the specified Filter</returns>
		public virtual List<Entry> FindUsersAndGroupsAndComputers(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&(|(objectClass=computer)(&(objectClass=Group)(objectCategory=Group))(&(objectClass=User)(objectCategory=Person)))({0}))", Filter);
			Searcher.Filter = Filter;
			return FindAll();
		}

		#endregion

		#region FindUsersAndGroupsAndComputersByName
		/// <summary>
		/// Finds a computer by computer name
		/// </summary>
		/// <param name="ComputerName">Computer name to search by</param>
		/// <returns>The computer's entry</returns>
		public virtual List<Entry> FindUsersAndGroupsAndComputersByName(string Name)
		{
			if (string.IsNullOrEmpty(Name))
				throw new ArgumentNullException("Name");
			Name = Name + "*";
			return FindUsersAndGroupsAndComputers("(|(name=" + Name + ")(sAMAccountName=" + Name + "))");
		}
		#endregion

		//#region FindActiveGroupMembers

		///// <summary>
		///// Returns a group's list of members
		///// </summary>
		///// <param name="GroupName">The group's name</param>
		///// <returns>A list of the members</returns>
		//public virtual List<Entry> FindActiveGroupMembers(string GroupName)
		//{
		//    try
		//    {
		//        List<Entry> Entries = this.FindGroups("cn=" + GroupName);
		//        return (Entries.Count < 1) ? new List<Entry>() : this.FindActiveUsersAndGroups("memberOf=" + Entries[0].DistinguishedName);
		//    }
		//    catch
		//    {
		//        return new List<Entry>();
		//    }
		//}

		//#endregion

		//#region FindActiveGroups

		///// <summary>
		///// Finds all active groups
		///// </summary>
		///// <param name="Filter">Filter used to modify the query</param>
		///// <param name="args">Additional arguments (used in string formatting</param>
		///// <returns>A list of all active groups' entries</returns>
		//public virtual List<Entry> FindActiveGroups(string Filter, params object[] args)
		//{
		//    Filter = string.Format(Filter, args);
		//    Filter = string.Format("(&((userAccountControl:1.2.840.113556.1.4.803:=512)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(cn=*$)))({0}))", Filter);
		//    return FindGroups(Filter);
		//}

		//#endregion

		#region FindActiveUsersAndGroups

		/// <summary>
		/// Finds all active users and groups
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all active groups' entries</returns>
		public virtual List<Entry> FindActiveUsersAndGroups(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&((userAccountControl:1.2.840.113556.1.4.803:=512)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(cn=*$)))({0}))", Filter);
			return FindUsersAndGroups(Filter);
		}

		#endregion

		#endregion Entry Functions

		#region User Functions

		#region FindUsers
		/// <summary>
		/// Finds all users
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all users meeting the specified Filter</returns>
		public virtual List<User> FindUsers(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&(objectClass=User)(objectCategory=Person)({0}))", Filter);
			Searcher.Filter = Filter;
			List<User> ReturnedResults = new List<User>();
			using (SearchResultCollection Results = Searcher.FindAll())
			{
				foreach (SearchResult Result in Results)
					ReturnedResults.Add(new User(Result.GetDirectoryEntry(), this.Context));
			}
			return ReturnedResults;
		}
		#endregion

		#region FindActiveUsers

		/// <summary>
		/// Finds all active users
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all active users' entries</returns>
		public virtual List<User> FindActiveUsers(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&((userAccountControl:1.2.840.113556.1.4.803:=512)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(cn=*$)))({0}))", Filter);
			return FindUsers(Filter);
		}

		#endregion

		#region FindUserByUserName
		/// <summary>
		/// Finds a user by his user name
		/// </summary>
		/// <param name="UserName">User name to search by</param>
		/// <returns>The user's entry</returns>
		public virtual User FindUserByUserName(string UserName)
		{
			if (string.IsNullOrEmpty(UserName))
				throw new ArgumentNullException("UserName");
			return FindUsers("sAMAccountName=" + UserName).FirstOrDefault();
		}
		#endregion

		#region FindUsersByName

		/// <summary>
		/// Finds users based on various fields related to the name similarly to an ADUC name search
		/// </summary>
		/// <param name="Name"></param>
		/// <returns></returns>
		public List<User> FindUsersByName(string Name)
		{
			if (string.IsNullOrEmpty(Name))
				throw new ArgumentNullException("Name");
			Name = Name + "*";
			
			return FindUsers("(|(sAMAccountName={0})(name={0})(givenname={0})(sn={0})(displayname={0}))", Name);
		}

		#endregion FindUsersByName

		#region FindUserPrinciples
		/// <summary>
		/// Finds all users
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all users meeting the specified Filter</returns>
		public virtual List<UserPrincipal> FindUserPrinciples(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&(objectClass=User)(objectCategory=Person)({0}))", Filter);
			Searcher.Filter = Filter;
			List<UserPrincipal> ReturnedResults = new List<UserPrincipal>();
			using (var context = new PrincipalContext(ContextType.Domain))
			using (SearchResultCollection Results = Searcher.FindAll())
			{
				foreach (SearchResult Result in Results)
					ReturnedResults.Add(UserPrincipal.FindByIdentity(context, IdentityType.DistinguishedName, Tools.GetDNFromPath(Result.Path)));
			}
			return ReturnedResults;
		}
		#endregion

		#region FindUserPrincipleByUserName

		/// <summary>
		/// Finds a user by his user name
		/// </summary>
		/// <param name="UserName">User name to search by</param>
		/// <returns>The user's entry</returns>
		public virtual UserPrincipal FindUserPrincipleByUserName(string UserName)
		{
			if (string.IsNullOrEmpty(UserName))
				throw new ArgumentNullException("UserName");
			return FindUserPrinciples("samAccountName=" + UserName).FirstOrDefault();
		}

		#endregion

		#region CreateUser

		/// <summary>
		/// Creates a user object
		/// </summary>
		/// <param name="CN"></param>
		/// <returns></returns>
		public User CreateUser(string Name)
		{
			DirectoryEntry newuser = Entry.Children.Add(Tools.GetCNFromName(Name), "user");
			return new User(newuser, this.Context);
		}

		#endregion CreateUser

		#endregion User Functions

		#region Group Functions

		#region FindGroups

		/// <summary>
		/// Finds all security groups
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all groups meeting the specified Filter</returns>
		public virtual List<Group> FindGroups(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&(objectClass=Group)(objectCategory=Group)({0}))", Filter);
			Searcher.Filter = Filter;
			List<Group> ReturnedResults = new List<Group>();
			using (SearchResultCollection Results = Searcher.FindAll())
			{
				foreach (SearchResult Result in Results)
					ReturnedResults.Add(new Group(Result.GetDirectoryEntry(), _Context));
			}
			return ReturnedResults;
		}

		#endregion FindGroups

		#region FindGroupByGroupName
		/// <summary>
		/// Finds a group by group name
		/// </summary>
		/// <param name="UserName">Group name to search by</param>
		/// <returns>The group entry</returns>
		public virtual Group FindGroupByGroupName(string GroupName)
		{
			if (string.IsNullOrEmpty(UserName))
				throw new ArgumentNullException("GroupName");
			return FindGroups("sAMAccountName=" + GroupName).FirstOrDefault();
		}
		#endregion FindGroupByGroupName

		#region AddUserToGroup
		/// <summary>
		/// Add a user to a security group
		/// </summary>
		/// <param name="UserName"></param>
		/// <param name="GroupName"></param>
		public void AddUserToGroup(string UserName, string GroupName)
		{
			using (Group group = FindGroupByGroupName(GroupName))
			{
				group.AddUser(UserName);
			}

		}

		#endregion AddUserToGroup

		#region RemoveUserFromGroup
		/// <summary>
		/// Add a user to a security group
		/// </summary>
		/// <param name="UserName">User to add to the group</param>
		/// <param name="GroupName">Group for the user to be added to</param>
		public void RemoveUserFromGroup(string UserName, string GroupName)
		{
			using (Group group = FindGroupByGroupName(GroupName))
			{
				group.RemoveUser(UserName);
			}				
		}

		#endregion RemoveUserFromGroup

		#region CreateGroup

		/// <summary>
		/// Creates a security group
		/// </summary>
		/// <param name="Name">Name of the security group</param>
		/// <returns></returns>
		public Group CreateGroup(string Name)
		{

			DirectoryEntry newgroup = Entry.Children.Add(Tools.GetCNFromName(Name), "group");						
			return new Group(newgroup, _Context);
		}

		#endregion CreateSecurityGroup

		#endregion Group Functions

		#region Computer Functions

		#region FindComputers

		/// <summary>
		/// Finds all computers
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all computers meeting the specified Filter</returns>
		public virtual List<Computer> FindComputers(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&(objectClass=computer)({0}))", Filter);
			Searcher.Filter = Filter;
			List<Computer> ReturnedResults = new List<Computer>();
			using (SearchResultCollection Results = Searcher.FindAll())
			{
				foreach (SearchResult Result in Results)
					ReturnedResults.Add(new Computer(Result.GetDirectoryEntry(), _Context));
			}
			return ReturnedResults;
		}

		#endregion FindComputers

		#region FindComputerByComputerName
		/// <summary>
		/// Finds a computer by computer name
		/// </summary>
		/// <param name="ComputerName">Computer name to search by</param>
		/// <returns>The computer's entry</returns>
		public virtual Computer FindComputerByComputerName(string ComputerName)
		{
			if (string.IsNullOrEmpty(ComputerName))
				throw new ArgumentNullException("ComputerName");
			return FindComputers("name=" + ComputerName).FirstOrDefault();
		}
		#endregion

		#region FindComputersByName
		/// <summary>
		/// Finds a computer by computer name
		/// </summary>
		/// <param name="ComputerName">Computer name to search by</param>
		/// <returns>The computer's entry</returns>
		public virtual List<Computer> FindComputersByName(string ComputerName)
		{
			if (string.IsNullOrEmpty(ComputerName))
				throw new ArgumentNullException("ComputerName");
			ComputerName = ComputerName + "*";
			return FindComputers("name=" + ComputerName);
		}
		#endregion

		#region CreateComputer

		/// <summary>
		/// Create a new computer object
		/// </summary>
		/// <param name="Name">Name of the computer</param>
		/// <returns>A computer entry</returns>
		public Computer CreateComputer(string Name)
		{
			if (FindComputerByComputerName(Name) != null)
				throw new Exception("There is already a computer with this name");
			DirectoryEntry newcomp = Entry.Children.Add(Tools.GetCNFromName(Name), "computer");
			return new Computer(newcomp, _Context);
		}

		#endregion CreateComputer

		#endregion Computer Functions

		#region Directory Functions

		#region FindOneLevelOUs

		/// <summary>
		/// Finds all ous
		/// </summary>
		/// <param name="Filter">Filter used to modify the query</param>
		/// <param name="args">Additional arguments (used in string formatting</param>
		/// <returns>A list of all computers meeting the specified Filter</returns>
		public List<Directory> FindOneLevelOUs(bool goUp)
		{
			if (goUp && _Entry.Parent != null && _Entry.Parent.Parent != null)
				_Entry = _Entry.Parent.Parent;
			else if (goUp)
				return null;
			//Filter = string.Format(Filter, args);
			string Filter = string.Format("(&(objectClass=organizationalunit))");
			Searcher.Filter = Filter;
			Searcher.SearchScope = SearchScope.OneLevel;
			var results = FindAll();
			Searcher.SearchScope = SearchScope.Subtree;
			List<Directory> ouresults = new List<Directory>();
			foreach (Entry e in results)
			{				
				ouresults.Add(new Directory("", this._UserName, this._Password, e.DirectoryEntry.Path, PropertiesToLoad));
			}
			return ouresults;
		}
		
		#endregion FindOneLevelOUs

		#region FindOUs

		/// <summary>
		/// Find OUs
		/// </summary>
		/// <param name="Filter"></param>
		/// <param name="args"></param>
		/// <returns></returns>
		public List<Directory> FindOUs(string Filter, params object[] args)
		{
			Filter = string.Format(Filter, args);
			Filter = string.Format("(&(objectClass=organizationalunit)({0}))", Filter);
			Searcher.Filter = Filter;			
			List<Directory> ReturnedResults = new List<Directory>();
			using (SearchResultCollection Results = Searcher.FindAll())
			{
				foreach (SearchResult Result in Results)
				{
					ReturnedResults.Add(new Directory("", this._UserName, this._Password, Result.Path, PropertiesToLoad));
				}
			}
			return ReturnedResults;
		}

		#endregion FindOUs

		#region FindAllOUs

		/// <summary>
		/// Returns all OUs under this directory
		/// </summary>
		/// <returns></returns>
		public List<Directory> FindAllOUs()
		{
			return FindOUs("name=*");
		}

		#endregion FindAllOUs

		#region FineSingleOU

		/// <summary>
		/// Find a single ou
		/// </summary>
		/// <param name="OUName">OU Name</param>
		/// <param name="goOneLevelDown">Search only one level</param>
		/// <returns></returns>
		public Directory FindSingleOU(string OUName, bool goOneLevelDown)
		{
			string Filter = string.Format("(&(objectClass=organizationalunit)(name={0}))", OUName);
			Searcher.Filter = Filter;
			if (goOneLevelDown)
				Searcher.SearchScope = SearchScope.OneLevel;
			var results = FindAll();
			Searcher.SearchScope = SearchScope.Subtree;
			if (results.Count > 1)
				throw new Exception("More than one OU has been returned");
			return (new Directory("", this._UserName, this._Password, results.First().DirectoryEntry.Path, PropertiesToLoad));
		}

		#endregion FineSingleOU

		#region FindOUByDistinguishedName

		/// <summary>
		/// Find a single ou
		/// </summary>
		/// <param name="OUName">OU Name</param>
		/// <param name="goOneLevelDown">Search only one level</param>
		/// <returns></returns>
		public Directory FindOUByDistinguishedName(string DistinguishedName)
		{
			string Filter = string.Format("distinguishedName={0}", DistinguishedName);
			return FindOUs(Filter).First();
		}

		#endregion FineSingleOU

		/// <summary>
		/// Gets a value from the entry
		/// </summary>
		/// <param name="Property">Property you want the information about</param>
		/// <returns>an object containing the property's information</returns>
		public virtual object GetValue(string Property)
		{
			PropertyValueCollection Collection = Entry.Properties[Property];
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
			PropertyValueCollection Collection = Entry.Properties[Property];
			return Collection != null ? Collection[Index] : null;
		}

		/// <summary>
		/// Sets a property of the entry to a specific value
		/// </summary>
		/// <param name="Property">Property of the entry to set</param>
		/// <param name="Value">Value to set the property to</param>
		public virtual void SetValue(string Property, object Value)
		{
			PropertyValueCollection Collection = Entry.Properties[Property];
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
			PropertyValueCollection Collection = Entry.Properties[Property];
			if (Collection != null)
				Collection[Index] = Value;
		}

		#endregion Directory Functions

		#region Domain Functions

		/// <summary>
		/// Gets a listing of domain controllers
		/// </summary>
		public List<string> GetDomainControllers()
		{
			DirectoryContext dcontext = new DirectoryContext(DirectoryContextType.DirectoryServer, Domain, this.UserName, this.Password);			
			List<string> results = new List<string>();
			foreach (DomainController dc in System.DirectoryServices.ActiveDirectory.Domain.GetDomain(dcontext).DomainControllers)
			{
				results.Add(dc.Name);
			}
			return results;
		}

		#endregion Domain Functions

		#endregion

		#region Properties

		/// <summary>
		/// The principal context for this directory
		/// </summary>
		public virtual PrincipalContext Context
		{
			get
			{
				return _Context;
			}
		}

		/// <summary>
		/// Domain used for this directory
		/// </summary>
		public virtual string Domain
		{
			get
			{
				return _Domain;
			}
		}

		/// <summary>
		/// Domain DN used for this directory
		/// </summary>
		public virtual string DomainDN
		{
			get
			{
				return _DomainDN;
			}
		}

		/// <summary>
		/// Path of the server
		/// </summary>
		public virtual string Path
		{
			get { return _Path; }
			set
			{
				_Path = value;
				if (Entry != null)
				{
					Entry.Close();
					Entry.Dispose();
					_Entry = null;
				}
				if (Searcher != null)
				{
					Searcher.Dispose();
					Searcher = null;
				}
				_Entry = new DirectoryEntry(_Path, _UserName, _Password, AuthenticationTypes.Secure);
				Searcher = new DirectorySearcher(Entry);
				Searcher.Filter = Query;
				Searcher.PageSize = 1000;
			}
		}

		/// <summary>
		/// User name used to log in
		/// </summary>
		public virtual string UserName
		{
			get { return _UserName; }
			set
			{
				_UserName = value;
				if (Entry != null)
				{
					_Entry.Close();
					_Entry.Dispose();
					_Entry = null;
				}
				if (Searcher != null)
				{
					Searcher.Dispose();
					Searcher = null;
				}
				_Entry = new DirectoryEntry(_Path, _UserName, _Password, AuthenticationTypes.Secure);
				Searcher = new DirectorySearcher(Entry);
				Searcher.Filter = Query;
				Searcher.PageSize = 1000;
			}
		}

		/// <summary>
		/// Password used to log in
		/// </summary>
		public virtual string Password
		{
			get { return _Password; }
			set
			{
				_Password = value;
				if (Entry != null)
				{
					_Entry.Close();
					_Entry.Dispose();
					_Entry = null;
				}
				if (Searcher != null)
				{
					Searcher.Dispose();
					Searcher = null;
				}
				_Entry = new DirectoryEntry(_Path, _UserName, _Password, AuthenticationTypes.Secure);
				Searcher = new DirectorySearcher(Entry);
				Searcher.Filter = Query;
				Searcher.PageSize = 1000;
			}
		}

		/// <summary>
		/// The query that is being made
		/// </summary>
		public virtual string Query
		{
			get { return _Query; }
			set
			{
				_Query = value;
				Searcher.Filter = _Query;
			}
		}

		/// <summary>
		/// Decides what to sort the information by
		/// </summary>
		public virtual string SortBy
		{
			get { return _SortBy; }
			set
			{
				_SortBy = value;
				Searcher.Sort.PropertyName = _SortBy;
				Searcher.Sort.Direction = SortDirection.Ascending;
			}
		}

		/// <summary>
		/// Determines if current directory has children
		/// </summary>
		public virtual bool HasChildren
		{
			get { return (Entry.Children != null); }
		}

		/// <summary>
		/// Determines if current directory has a higher level of ou's
		/// </summary>
		public virtual bool HasParent
		{
			get { return (Entry.Parent != null && Entry.Parent.Parent != null); }
		}

		/// <summary>
		/// The Parent Path for this entry
		/// </summary>
		public virtual string ParentPath
		{
			get { return Entry.Parent.Path; }
		}

		/// <summary>
		/// the search scope for all queries
		/// </summary>
		public virtual bool SearchOneLevel
		{
			get { return Searcher.SearchScope == SearchScope.OneLevel; }
			set { if (value) Searcher.SearchScope = SearchScope.OneLevel; else Searcher.SearchScope = SearchScope.Subtree; }
		}

		/// <summary>
		/// Get or set the properties to load during a search
		/// </summary>
		public virtual string[] PropertiesToLoad { get;	set;}

		/// <summary>
		/// the name attribute for this directory's entry
		/// </summary>
		public virtual string Name
		{
			get { return (string)GetValue("name"); }
			set { SetValue("name", value); }
		}

		/// <summary>
		/// the OU attribute for this directory's entry
		/// </summary>
		public virtual string OuName
		{
			get { return (string)GetValue("ou"); }
			set { SetValue("ou", value); }
		}

		/// <summary>
		/// Limit the number of results to return.  0 returns all
		/// </summary>
		public virtual int SizeLimit
		{
			get { return Searcher.SizeLimit; }
			set { Searcher.SizeLimit = value; }
		}

		/// <summary>
		/// The directory entry for this directory
		/// </summary>
		public virtual DirectoryEntry Entry
		{
			get { return _Entry; }
		}

		#endregion

		#region Private Variables
		private string _Path = "";
		private string _UserName = "";
		private string _Password = "";
		private DirectoryEntry _Entry = null;
		private string _Query = "";
		private DirectorySearcher Searcher = null;
		private string _SortBy = "";
		private PrincipalContext _Context;
		private string _Domain;
		private string _DomainDN;
		#endregion

		#region IDisposable Members

		public void Dispose()
		{
			if (_Entry != null)
			{
				_Entry.Close();
				_Entry.Dispose();
				_Entry = null;
			}
			if (Searcher != null)
			{
				Searcher.Dispose();
				Searcher = null;
			}
		}
			
		#endregion			
	}
}