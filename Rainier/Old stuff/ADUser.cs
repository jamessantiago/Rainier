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
    public partial class ADUser
    {
        public SelectList Clearances { get; set; }
        public SelectList Units { get; set; }
        public SelectList Locations { get; set; }
        public SelectList Branches { get; set; }

        public long UnitID { get; set; }
        public long LocationID { get; set; }
        public long BranchID { get; set; }
        public string OUPath { get; set; }
        public string OUName { get; set; }
        public long OriginalOUID { get; set; }
        public string OriginalOUPath { get; set; }
        public string OriginalOUName { get; set; }

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
					IQueryable<IASO2OU> iaso2ous = _db.IASO2OUs.Where(x => x.OUID == parentid && x.IASO.IsActive && x.IASO.Expires >= DateTime.Now);
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

        public string SSNPrivate
        {
            get
            {
                string ssnprivate = "";
                if (!string.IsNullOrEmpty(this.SSN))
                {
                    char[] ssn = this.SSN.ToCharArray();
                    for (int i = 0; i < ssn.Length; i++)
                    {
                        if (i > ssn.Length - 5) ssnprivate += ssn[i];
                        else ssnprivate += "*";
                    }
                }
                return ssnprivate;
            }
        }

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

    }
}
