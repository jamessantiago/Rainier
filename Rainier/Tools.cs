using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using ActiveDs;
using System.DirectoryServices.ActiveDirectory;

namespace Rainier
{	
	/// <summary>
	/// Functions needed for active directory
	/// </summary>
	public static class Tools
	{
		#region Conversion Tools

		/// <summary>
		/// Returns a CN
		/// </summary>
		/// <param name="Name"></param>
		/// <returns>CN</returns>
		public static string GetCNFromName(string Name)
		{
			return "CN=" + Name.Replace(",", "\\,");
		}

		/// <summary>
		/// Returns a DN
		/// </summary>
		/// <param name="Path">The path from a directoryentry object</param>
		/// <returns>DN</returns>
		public static string GetDNFromPath(string Path)
		{
			Path = Path.Remove(0, Path.Contains("OU") ? Path.IndexOf("OU") : Path.IndexOf("DC"));
			return Path;
		}

		/// <summary>
		/// LDAP Path
		/// </summary>
		/// <param name="DomainName">Fully qualified domain name</param>
		/// <param name="Server">Domain Controller or FQDN</param>
		/// <returns>LDAP Path</returns>
		public static string GetRootLDAPPathFromDomainName(string DomainName, string Server)
		{
			DomainName = "/DC=" + DomainName.Replace(".", ",DC=");
			return String.Format("LDAP://{0}{1}", Server, DomainName);
		}

		/// <summary>
		/// return[0] is the server return[1] is the domain
		/// </summary>
		/// <param name="Path"></param>
		/// <returns>return[0] is the server return[1] is the domain</returns>
		public static string[] ReverseRootLDAPPath(string Path)
		{
		    Path = Path.Replace("LDAP://", "");
		    string[] result = new string[2];
		    result[0] = Path.Substring(0, Path.IndexOf("/") != -1 ? Path.IndexOf("/") : Path.Length);
			string removethis = Path.Substring(0, Path.IndexOf("DC="));
			result[1] = Path.Replace(removethis, "");
			return result;
		}

		/// <summary>
		/// the parent path from an LDAP path string
		/// </summary>
		/// <param name="Path">Full path of the child item</param>
		/// <param name="Name">Name of the OU</param>
		/// <returns></returns>
		public static string GetParentFromPath(string Path, string Name)
		{
			Name = "OU=" + Name.Replace(@",", @"\,") + ",";
			return Path.Replace(Name, "");
		}

		/// <summary>
		/// Determines if an ldap path is teh child of another ldap path
		/// </summary>
		/// <param name="Path"></param>
		/// <param name="ParantPath"></param>
		/// <returns></returns>
		public static bool IsPathChildOfPath(string Path, string ParantPath)
		{
			string child = GetDNFromPath(Path).ToUpper();
			string parant = GetDNFromPath(ParantPath).ToUpper();
			bool results = false;
			try
			{
				results = (child.Contains(parant));
			}
			catch {  }
			return results;
		}

		/// <summary>
		/// Converts active directory's large integer type to a datetime
		/// </summary>
		/// <param name="filetime"></param>
		/// <returns></returns>
		public static DateTime ToDateTime(LargeInteger filetime)
		{
			if (filetime == null)
				return DateTime.MinValue;
			long longDate = (((long)(filetime.HighPart) << 32) + (long)filetime.LowPart);
			return DateTime.FromFileTime(longDate);
		}

		/// <summary>
		/// Converts a datetime to active directory's large integer type
		/// </summary>
		/// <param name="datetime"></param>
		/// <returns></returns>
		public static LargeInteger ToLargeInteger(DateTime datetime)
		{
			Int64 filetime = datetime.ToFileTime();
			LargeInteger largeinteger = new LargeInteger();
			largeinteger.HighPart = (int)(filetime >> 32);
			largeinteger.LowPart = (int)(filetime & 0xFFFFFFFF);
			return largeinteger;
		}

		/// <summary>
		/// Converts a datetime to active directory's large integer type
		/// </summary>
		/// <param name="datetime"></param>
		/// <returns></returns>
		public static string ToLargeIntegerString(DateTime datetime)
		{
			Int64 filetime = datetime.ToFileTime();
			LargeInteger largeinteger = new LargeInteger();
			largeinteger.HighPart = (int)(filetime >> 32);
			largeinteger.LowPart = (int)(filetime & 0xFFFFFFFF);
			return Math.Abs(largeinteger.HighPart).ToString() + Math.Abs(largeinteger.LowPart).ToString();
		}

		public static string ToUtcCodedTime(DateTime datetime)
		{
			return datetime.ToString("yyyyMMddHHmmss.0z");
		}

		#endregion Conversion Tools

		#region Domain Tools

		/// <summary>
		/// Gets the current domain name
		/// </summary>
		/// <returns></returns>
		public static string GetCurrentDomain()
		{
			return Domain.GetCurrentDomain().Name;
		}

		/// <summary>
		/// Gets a listing of domain controllers
		/// </summary>
		public static List<string> GetDomainControllers(string DomainName)
		{
			DirectoryContext context = new DirectoryContext(DirectoryContextType.Domain, DomainName);
			List<string> results = new List<string>();
			foreach (DomainController dc in Domain.GetDomain(context).DomainControllers)
			{				
				results.Add(dc.Name);
			}
			return results;
		}

		#endregion Domain Tools

		#region AD Properties

		/// <summary>
		/// GUID values for ACE object types
		/// </summary>
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

		/// <summary>
		/// aces for giving permisions to join a computer to a domain
		/// </summary>
		public class ADACEComputerJoinPermissions
		{

			/// <summary>
			/// Constructor
			/// </summary>
			/// <param name="trustee">Trustee to assign permissions to.  Defaults to Authenticated Users</param>
			public ADACEComputerJoinPermissions(string trustee)
			{
				if (!string.IsNullOrEmpty(trustee))
					Trustee = trustee;
				else
					Trustee = "Authenticated Users";
			}

			/// <summary>
			/// The trustee for the permissions, defaults to authenticated users
			/// </summary>
			public virtual string Trustee { get; set; }			

			/// <summary>
			/// List of aces
			/// </summary>
			public virtual List<IADsAccessControlEntry> ace_writeaccountrestrictions {
				get {
					List<IADsAccessControlEntry> list = new List<IADsAccessControlEntry>();
					IADsAccessControlEntry ace_writeaccountrestrictions = new AccessControlEntry();
					ace_writeaccountrestrictions.Trustee = Trustee;
					ace_writeaccountrestrictions.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
					ace_writeaccountrestrictions.AceFlags = 0;
					ace_writeaccountrestrictions.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_writeaccountrestrictions.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_writeaccountrestrictions.ObjectType = Tools.ADACEObjectTypes.USER_ACCOUNT_RESTRICTIONS;
					IADsAccessControlEntry ace_validatedspn = new AccessControlEntry();
					ace_validatedspn.Trustee = Trustee;
					ace_validatedspn.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_SELF);
					ace_validatedspn.AceFlags = 0;
					ace_validatedspn.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;				
					ace_validatedspn.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_validatedspn.ObjectType = Tools.ADACEObjectTypes.VALIDATED_SPN;
					IADsAccessControlEntry ace_validateddns = new AccessControlEntry();
					ace_validateddns.Trustee = Trustee;
					ace_validateddns.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_SELF);
					ace_validateddns.AceFlags = 0;
					ace_validateddns.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_validateddns.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_validateddns.ObjectType = Tools.ADACEObjectTypes.VALIDATED_DNS_HOST_NAME;
					IADsAccessControlEntry ace_resetpw = new AccessControlEntry();
					ace_resetpw.Trustee = Trustee;
					ace_resetpw.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
					ace_resetpw.AceFlags = 0;
					ace_resetpw.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_resetpw.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_resetpw.ObjectType = Tools.ADACEObjectTypes.RESET_PASSWORD_GUID;
					IADsAccessControlEntry ace_changepw = new AccessControlEntry();
					ace_changepw.Trustee = Trustee;
					ace_changepw.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
					ace_changepw.AceFlags = 0;
					ace_changepw.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_changepw.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_changepw.ObjectType = Tools.ADACEObjectTypes.USER_CHANGE_PASSWORD;
					IADsAccessControlEntry ace_receiveas = new AccessControlEntry();
					ace_receiveas.Trustee = Trustee;
					ace_receiveas.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
					ace_receiveas.AceFlags = 0;
					ace_receiveas.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_receiveas.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_receiveas.ObjectType = Tools.ADACEObjectTypes.RECEIVE_AS;
					IADsAccessControlEntry ace_sendas = new AccessControlEntry();
					ace_sendas.Trustee = Trustee;
					ace_sendas.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
					ace_sendas.AceFlags = 0;
					ace_sendas.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_sendas.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_sendas.ObjectType = Tools.ADACEObjectTypes.SEND_AS;
					IADsAccessControlEntry ace_access = new AccessControlEntry();
					ace_access.Trustee = Trustee;
					ace_access.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS |
													ADS_RIGHTS_ENUM.ADS_RIGHT_ACTRL_DS_LIST |
													ADS_RIGHTS_ENUM.ADS_RIGHT_GENERIC_READ |
													ADS_RIGHTS_ENUM.ADS_RIGHT_DELETE |
													ADS_RIGHTS_ENUM.ADS_RIGHT_DS_DELETE_TREE |
													ADS_RIGHTS_ENUM.ADS_RIGHT_READ_CONTROL);
					ace_access.AceFlags = 0;
					ace_access.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED;
					IADsAccessControlEntry ace_delete = new AccessControlEntry();
					ace_delete.Trustee = Trustee;
					ace_delete.AccessMask = 197076;// (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DELETE | ADS_RIGHTS_ENUM.ADS_RIGHT_DS_DELETE_TREE);
					ace_delete.AceFlags = 0;
					ace_delete.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_delete.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED;
					IADsAccessControlEntry ace_description = new AccessControlEntry();
					ace_description.Trustee = Trustee;
					ace_description.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
					ace_description.AceFlags = 0;
					ace_description.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_description.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_description.ObjectType = Tools.ADACEObjectTypes.DESCRIPTION_GUID;
					IADsAccessControlEntry ace_logoninfo = new AccessControlEntry();
					ace_logoninfo.Trustee = Trustee;
					ace_logoninfo.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
					ace_logoninfo.AceFlags = 0;
					ace_logoninfo.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_logoninfo.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_logoninfo.ObjectType = Tools.ADACEObjectTypes.USER_LOGON_INFORMATION;
					IADsAccessControlEntry ace_cname = new AccessControlEntry();
					ace_cname.Trustee = Trustee;
					ace_cname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
					ace_cname.AceFlags = 0;
					ace_cname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_cname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_cname.ObjectType = Tools.ADACEObjectTypes.COMMONNAME_GUID;
					IADsAccessControlEntry ace_dname = new AccessControlEntry();
					ace_dname.Trustee = Trustee;
					ace_dname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
					ace_dname.AceFlags = 0;
					ace_dname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_dname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_dname.ObjectType = Tools.ADACEObjectTypes.DISPLAYNAME_GUID;
					IADsAccessControlEntry ace_sname = new AccessControlEntry();
					ace_sname.Trustee = Trustee;
					ace_sname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_WRITE_PROP);
					ace_sname.AceFlags = 0;
					ace_sname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_sname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_sname.ObjectType = Tools.ADACEObjectTypes.SAMACCOUNTNAME_GUID;
					IADsAccessControlEntry ace_aaname = new AccessControlEntry();
					ace_aaname.Trustee = Trustee;
					ace_aaname.AccessMask = (int)(ADS_RIGHTS_ENUM.ADS_RIGHT_DS_CONTROL_ACCESS);
					ace_aaname.AceFlags = 0;
					ace_aaname.Flags = 3; // (int)ADS_FLAGTYPE_ENUM.ADS_FLAG_OBJECT_TYPE_PRESENT;
					ace_aaname.AceType = (int)ADS_ACETYPE_ENUM.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT;
					ace_aaname.ObjectType = Tools.ADACEObjectTypes.ALLOWED_TO_AUTHENTICATE;

					list.Add(ace_access);
					list.Add(ace_description);
					list.Add(ace_sname);
					list.Add(ace_dname);
					list.Add(ace_logoninfo);
					list.Add(ace_writeaccountrestrictions);
					list.Add(ace_validateddns);
					list.Add(ace_validatedspn);
					list.Add(ace_aaname);
					list.Add(ace_receiveas);
					list.Add(ace_sendas);
					list.Add(ace_changepw);
					list.Add(ace_resetpw);
					list.Add(ace_delete);
					list.Add(ace_cname);

					return list;
				}
			}
		}

		#endregion AD Properties
	}
}
