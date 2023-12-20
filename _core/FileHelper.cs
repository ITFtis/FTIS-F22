using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using static HRM_F22.Code;

namespace HRM_F22
{
    public class FileHelper
    {
        public static string GetFileFolder(TempUploadFile en)
        {
            string result = "";

            result = AppConfig.RootPath + "Temp/" + en.ToString() + "/";

            return result;
        }

        public static string GetFileFolder(UploadFile en)
        {
            string result = "";

            result = AppConfig.RootPath + en.ToString() + "/";

            return result;
        }
    }
}