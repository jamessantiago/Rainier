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
    public partial class OU
    {
        public string NiceName
        {
            get
            {
                if (this.Name != null)
                {
                    return this.Name.Substring(this.Name.LastIndexOf("=") + 1);
                }
                else return "";
            }
        }

        public string NicePath
        {
            get
            {
                if (this.Path != null)
                {
                    return this.Path.Substring(this.Path.LastIndexOf("/") + 1);
                }
                else return "";
            }
        }

        public string LDAPPath
        {
            get
            {
                if (this.Path != null)
                {
                    return this.Path.Substring(0, this.Path.LastIndexOf("/") + 1);
                }
                else return "";
            }
        }

        public List<OU> subOUs
        {
            get
            {
                ADManagerDataContext _db = new ADManagerDataContext();
                return _db.OUs.Where(x => x.ParentID == this.OUID && x.IsActive == true && x.IsDisplay == true).OrderBy(x => x.Name).ToList();
            }
        }

        public List<OU> subOUsAll
        {
            get
            {
                ADManagerDataContext _db = new ADManagerDataContext();
                return _db.OUs.Where(x => x.ParentID == this.OUID).ToList();
            }
        }

        public OU Parent
        {
            get
            {
                ADManagerDataContext _db = new ADManagerDataContext();
                OU parent = _db.OUs.SingleOrDefault(x => x.OUID == this.ParentID);
                if (parent == null) parent = new OU();
                return parent;
            }
        }

        public List<OU> subTree
        {
            get
            {
                List<OU> answer = new List<OU>();
                ADManagerDataContext _db = new ADManagerDataContext();
                answer.Add(this);
                List<OU> subOUs = this.subOUs;
                foreach (OU subou in subOUs)
                {
                    List<OU> subsubtree = subou.subTree;
                    foreach (OU subsubou in subsubtree)
                    {
                        answer.Add(subsubou);
                    }
                }
                return answer;
            }
        }

        public List<OU> subTreeAll
        {
            get
            {
                List<OU> answer = new List<OU>();
                ADManagerDataContext _db = new ADManagerDataContext();
                answer.Add(this);
                List<OU> subOUs = this.subOUsAll;
                foreach (OU subou in subOUs)
                {
                    List<OU> subsubtree = subou.subTreeAll;
                    foreach (OU subsubou in subsubtree)
                    {
                        answer.Add(subsubou);
                    }
                }
                return answer;
            }
        }

        public string subOUHTML
        {
            get
            {
                string answer = "";
                ADManagerDataContext _db = new ADManagerDataContext();
                IQueryable<OU> subOUs = _db.OUs.Where(x => x.ParentID == this.OUID && x.IsActive == true && x.IsDisplay == true).OrderBy(x => x.Name);
                if (subOUs.Count() > 1)
                {
                    answer += "<li><span class=\"folder\">" + this.NiceName + "</span>";
                    foreach (OU subOU in subOUs)
                    {
                        string hidestring = "";
                        if (subOU.OULevel > 1) hidestring = " style=\"display:none;\"";
                        answer += "<ul id=\"OU" + subOU.OUID + "\"" + hidestring + ">";
                        answer += subOU.subOUHTML;
                        answer += "</ul>";
                    }
                    answer += "</li>";
                }
                else
                {
                    answer = "<li><span class=\"file\"><a id=\"link" + this.OUID + "\" onclick=\"loadUsers(" + this.OUID + ", false)\">" + this.NiceName + "</a></span></li>";
                }
                return answer;
            }
        }

        public bool IsUserHome
        {
            get
            {
                bool answer = false;
                ADManagerDataContext _db = new ADManagerDataContext();
                IQueryable<OU> subous = _db.OUs.Where(x => x.ParentID == this.OUID && x.IsActive == true && x.Name.Equals("OU=Users"));
                if (subous != null && subous.Count() > 0) answer = true;
                return answer;
            }
        }

		public bool IsGroupHome
		{
			get
			{
				bool answer = false;
				ADManagerDataContext _db = new ADManagerDataContext();
				IQueryable<OU> subous = _db.OUs.Where(x => x.ParentID == this.OUID && x.IsActive == true && x.Name.Equals("OU=Groups"));
				if (subous != null && subous.Count() > 0) answer = true;
				return answer;
			}
		}

		public bool HasSubOUs
		{
			get
			{
				bool answer = false;
				ADManagerDataContext _db = new ADManagerDataContext();
				IQueryable<OU> subOUs = _db.OUs.Where(x => x.ParentID == this.OUID && x.IsActive == true && x.IsDisplay == true).OrderBy(x => x.Name);
				if (subOUs.Count() > 1) answer = true;
				return answer;
			}
		}

        public List<ADUser> IASOs
        {
            get
            {
                ADManagerDataContext _db = new ADManagerDataContext();
                List<long> parentids = new List<long>();
                List<ADUser> iasos = new List<ADUser>();
                OU currparent = this.Parent;
                while (currparent != null && currparent.OUID > 0)
                {
                    parentids.Add(currparent.OUID);
                    currparent = currparent.Parent;
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
						catch (Exception e)
						{
							EMailUtil.EmailMike("Error in FInding IASO", e.ToString());
						}
                    }
                }
                return iasos;
            }
        }

    }
}
