using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HRM_F22.Controllers.Path
{
    [Dou.Misc.Attr.MenuDef(Id = "HrPath", Name = "人資專區", Index = 1, IsOnlyPath = true)]
    public class HrPathController : Controller
    {
        // GET: HrPath
        public ActionResult Index()
        {
            return View();
        }
    }
}