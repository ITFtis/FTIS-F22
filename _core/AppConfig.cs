using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace HRM_F22
{
    public class AppConfig
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region 私有變數

        private static string _rootPath;

        #endregion

        #region 建構子

        static AppConfig()
        {
            try
            {
                _rootPath = ConfigurationManager.AppSettings["RootPath"].ToString();
            }
            catch (Exception ex)
            {
                logger.Info(ex.Message);
                logger.Info(ex.StackTrace);

                throw ex;
            }
        }

        #endregion

        #region 公用屬性      

        /// <summary>
        /// 檔案存放跟目錄
        /// </summary>
        public static string RootPath
        {
            get { return _rootPath; }
        }

        #endregion
    }
}