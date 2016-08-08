using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace Rainier
{
	/// <summary>
	/// An OU
	/// </summary>
	public class OU : Entry
	{
		#region Constructor
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="DirectoryEntry"></param>
		public OU(DirectoryEntry DirectoryEntry, PrincipalContext Context)
			: base(DirectoryEntry, Context)
		{
			this.DirectoryEntry = DirectoryEntry;
		}
		#endregion Constructor
	}
}
