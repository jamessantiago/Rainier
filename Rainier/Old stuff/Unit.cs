using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.Linq;
using ADManager.Controllers;

namespace ADManager.Models
{
    public partial class Unit : IComparable
    {
        public int CompareTo(object obj)
        {
            Unit Compare = (Unit)obj;
            int result = this.UnitName.CompareTo(Compare.UnitName);
            if (result == 0)
            {
                result = this.UnitName.CompareTo(Compare.UnitName);
            }
            return result;
        }
    }
}
