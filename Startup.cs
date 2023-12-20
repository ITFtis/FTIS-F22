using Microsoft.Owin;
using Owin;
using System;
using System.Threading.Tasks;

[assembly: OwinStartup(typeof(HRM_F22.Startup))]

namespace HRM_F22
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            Dou.Context.Init(new Dou.DouConfig
            {
                DefaultAdminUserId = "ftisadmin",
                DefaultPassword = "1230180",                
                PasswordEncode = (p) =>
                {
                    return (int.Parse(p) * 4 + 13579) + "";
                    //return System.Web.Helpers.Crypto.HashPassword(p);
                },
                VerifyPassword = (ep, vp) =>
                {
                    int pint = 0;
                    bool tp = int.TryParse(vp, out pint);
                    if (!tp)
                        return false;
                    else
                    {
                        return ep == (pint * 4 + 13579) + "";
                    }
                    //return System.Web.Helpers.Crypto.VerifyHashedPassword(ep, vp);
                },
                SqlDebugLog = true,
                LoginPage = new System.Web.Mvc.UrlHelper(System.Web.HttpContext.Current.Request.RequestContext).Action("DouLogin", "User")
            });
        }
    }
}
