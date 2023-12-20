using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HRM_F22.Controllers.Path
{
    [Dou.Misc.Attr.MenuDef(Id = "SysPath", Name = "系統管理", Index = int.MaxValue, IsOnlyPath = true)]
    public class SysPathController : Controller
    {
        // GET: SysPath
        public ActionResult Index()
        {
            return View();
        }
    }
}