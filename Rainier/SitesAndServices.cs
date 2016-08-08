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
	public class SitesAndServices
	{
		#region constructor

		public SitesAndServices(string UserName, string Password, string Path)
		{
			this.Path = Path;
			this._username = UserName;
			this._password = Password;	
			string[] serverinfo = Tools.ReverseRootLDAPPath(this.Path);
			Domain = serverinfo[0];
			this._context = new DirectoryContext(DirectoryContextType.DirectoryServer,  Domain, this._username, this._password);
			Forest.GetForest(_context);
		}

		#endregion constructor

		#region Properties

		public string Path = "";
		public string Domain = "";

		#endregion Properties

		#region Private Properties

		private string _username = "";
		private string _password = "";
		private DirectoryContext _context;

		#endregion Private Properties

		#region Public Properties

		public void CreateNewSubnet(string Name, ActiveDirectorySite Site)
		{
			ActiveDirectorySubnet newsub = new ActiveDirectorySubnet(this._context, Name, Site.Name);
			newsub.Save();
		}

		public void DeleteSubnet(string Name)
		{
			var subnet = FindSubnet(Name);
			subnet.Delete();
		}

		public List<ActiveDirectorySite> GetAllSites()
		{
			List<ActiveDirectorySite> sites = new List<ActiveDirectorySite>();
			foreach (ActiveDirectorySite site in Forest.GetForest(_context).Sites)
			{
				sites.Add(site);
			}
			return sites;
		}

		public ActiveDirectorySite FindSite(string Name)
		{
			var site = ActiveDirectorySite.FindByName(_context, Name);
			return site;
		}

		public ActiveDirectorySubnet FindSubnet(string Name)
		{

			var subnet = ActiveDirectorySubnet.FindByName(_context, Name);
			return subnet;
		}

		public List<ActiveDirectorySubnet> GetAllSubnets()
		{
			List<ActiveDirectorySubnet> subnets = new List<ActiveDirectorySubnet>();
			foreach (ActiveDirectorySite site in GetAllSites())
			{
				foreach (ActiveDirectorySubnet sub in site.Subnets)
				{
					subnets.Add(sub);
				}
			}
			return subnets;
		}

		#endregion Public Properties
	}
}
