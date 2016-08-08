using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.Linq;
using ADManager.Controllers;
using ADManager.Util;

namespace ADManager.Models
{
    public partial class Workstation
    {

		public string OUPath { get; set; }
        public string OUName { get; set; }

		public string Network { get; set; }
		public string Location { get; set; }
		public string Building { get; set; }
		public string Tunnel { get; set; }
        public string Other { get; set; }
		public string ComputerType { get; set; }
		public string UnitCode { get; set; }
        public string OtherUnit { get; set; }
        public string isCritical { get; set; }
		public bool Disabled { get; set; }

        public OU myOU
        {
            get
            {
                ADManagerDataContext _db = new ADManagerDataContext();
                OU userou = _db.OUs.SingleOrDefault(x => x.OUID == this.OUID);
                if (userou != null) return _db.OUs.SingleOrDefault(x => x.OUID == userou.ParentID);
                else return null;
            }
        }

		public List<ADUser> MyIASOs
		{
			get
			{
				List<ADUser> iasos = new List<ADUser>();
				ADManagerDataContext _db = new ADManagerDataContext();
				List<long> parentids = new List<long>();
				OU ou = _db.OUs.SingleOrDefault(x => x.Path.Equals(this.OUPath));
				if (ou != null)
				{
					parentids.Add(ou.OUID);
					OU currparent = ou.Parent;
					while (currparent != null && currparent.OUID > 0)
					{
						parentids.Add(currparent.OUID);
						currparent = currparent.Parent;
					}
				}
				foreach (long parentid in parentids)
				{
					IQueryable<IASO2OU> iaso2ous = _db.IASO2OUs.Where(x => x.OUID == parentid && x.IASO.IsActive);
					foreach (IASO2OU iaso2ou in iaso2ous)
					{
						try
						{
							ADUser iaso = ADHelper.GetUser(iaso2ou.IASO.Username);
							if (iaso != null)
							{
								iasos.Add(iaso);
							}
							else
							{

							}
						}
						catch (Exception e) { e.ToString(); }
					}
				}
				return iasos;
			}
		}

    }
}
