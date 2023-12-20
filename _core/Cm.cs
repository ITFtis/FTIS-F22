using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace HRM_F22
{
    public class Cm
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// 實體轉相對路徑
        /// </summary>
        /// <param name="imagesurl1"></param>
        /// <returns></returns>
        public static string PhysicalToUrl(string imagesurl1)
        {

            string tmpRootDir = HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath.ToString());//獲取程式根目錄

            string imagesurl2 = imagesurl1.Replace(tmpRootDir, ""); //轉換成相對路徑

            imagesurl2 = imagesurl2.Replace(@"/", @"/");

            //imagesurl2 = imagesurl2.Replace(@"Aspx_Uc/", @"");

            return imagesurl2;

        }

        /// <summary>
        /// 相對路徑轉實體
        /// </summary>
        /// <param name="imagesurl1"></param>
        /// <returns></returns>
        public static string UrlToPhysical(string imagesurl1)
        {

            string tmpRootDir = HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath.ToString());//獲取程式根目錄

            string imagesurl2 = tmpRootDir + imagesurl1.Replace(@"/", @"/"); //轉換成絕對路徑

            return imagesurl2;

        }
    }
}