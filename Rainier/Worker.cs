using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rainier;

///
namespace Rainier
{
	class Worker : IDisposable
	{
		//public Worker(string username, string password, string DomainName, string Server, bool loadNoneOrAllProperties)
		//{
		//    UserName = username;
		//    Password = password;
		//    RootLDAPPath = Tools.GetRootLDAPPathFromDomainName(DomainName, Server);
		//    RootDirectory = new Directory("", UserName, Password, RootLDAPPath, loadNoneOrAllProperties);			
		//}

		//public Worker(string username, string password, string DomainName, string Server, string[] PropertiesToLoad)
		//{
		//    UserName = username;
		//    Password = password;
		//    RootLDAPPath = Tools.GetRootLDAPPathFromDomainName(DomainName, Server);
		//    RootDirectory = new Directory("", UserName, Password, RootLDAPPath, PropertiesToLoad);
		//}

		#region Properties
		/// <summary>
		/// Username to connect to active directory with
		/// </summary>
		public string UserName {get; private set;}
		
		/// <summary>
		/// Password to connect to active directory with
		/// </summary>
		public virtual string Password {get; private set;}

		/// <summary>
		/// Root directory for the domain
		/// </summary>
		public Directory RootDirectory {get; private set;}

		/// <summary>
		/// Root LDAP Path
		/// </summary>
		public virtual string RootLDAPPath { get; private set; }

		#endregion Properties

		#region Private Properties

		#endregion Private Properties

		#region dispose

		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		~Worker()
		{
			Dispose(false);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (disposing)
			{
				//dispose
			}
		}

		#endregion dispose
	}

	//public class DirectoryLazyList<Directory> : List<Directory>
	//{
	//    public int TotalCount { get; private set; }

	//    public DirectoryLazyList(List<Directory> Source, Rainier.Directory currentOU)
	//    {
	//        TotalCount = Source.Count();
			
			
	//    }

	//    //public bool HasHigherLevel
	//    //{ get { return 
	//}

	
}
