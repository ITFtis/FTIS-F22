using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace HRM_F22
{
    public class DateFormat
    {
        /// <summary>
        /// 字串轉日期
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static DateTime ToDateTime(string date)
        {
            DateTime result = DateTime.MinValue;

            DateTime.TryParse(date, out result);

            return result;
        }

        /// <summary>
        /// yyyy-MM-dd
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate1(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToDate1(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// yyyy-MM-dd
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate1(DateTime date)
        {
            string result = "";

            result = date.ToString("yyyy-MM-dd");

            return result;
        }

        /// <summary>
        /// yyyy-MM-dd HH:mm:ss
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate2(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToDate2(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// yyyy-MM-dd HH:mm:ss
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate2(DateTime date)
        {
            string result = "";

            result = date.ToString("yyyy-MM-dd HH:mm:ss");

            return result;
        }

        /// <summary>
        /// MM
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate3(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToDate3(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// MM
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate3(DateTime date)
        {
            string result = "";

            result = date.ToString("MM");

            return result;
        }

        /// <summary>
        /// yyyy/MM/dd
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate4(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToDate4(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// yyyy/MM/dd
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate4(DateTime date)
        {
            string result = "";

            //當 datetime 為最小值時，回傳「空字串」。
            if (date == DateTime.MinValue)
                result = "";
            else
                result = date.ToString("yyyy/MM/dd");

            return result;
        }

        /// <summary>
        /// yyyy/MM
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate5(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToDate5(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// yyyy/MM
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate5(DateTime date)
        {
            string result = "";

            result = date.ToString("yyyy/MM");

            return result;
        }

        /// <summary>
        /// yyyy-MM-dd HH:mm
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate6(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToDate6(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// yyyy-MM-dd HH:mm
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate6(DateTime date)
        {
            string result = "";

            result = date.ToString("yyyy-MM-dd HH:mm");

            return result;
        }

        /// <summary>
        /// yyyy/MM/dd HH:mm
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate7(object date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date.ToString());
                result = string.Format("{0:yyyy/MM/dd HH:mm}", dd);
            }
            catch
            {
                return null;
            }

            return result;
        }

        /// <summary>
        /// yyyyMMdd
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate8(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToDate4(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// yyyyMMdd
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToDate8(DateTime date)
        {
            string result = "";

            //當 datetime 為最小值時，回傳「空字串」。
            if (date == DateTime.MinValue)
                result = "";
            else
                result = date.ToString("yyyyMMdd");

            return result;
        }

        /// <summary>
        /// 西元轉民國：2004-12-20 00:00:00.000 => 091/01/01
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToTwDate1(string date)
        {
            string result = "";

            try
            {
                DateTime dd = DateTime.Parse(date);
                result = ToTwDate1(dd);
            }
            catch
            {
                return date;
            }

            return result;
        }

        /// <summary>
        /// 西元轉民國：2004-12-20 00:00:00.000 => 091/01/01
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string ToTwDate1(DateTime date)
        {
            string result = "";

            CultureInfo culture = new CultureInfo("zh-TW");
            culture.DateTimeFormat.Calendar = new TaiwanCalendar();
            result = date.ToString("yyy", culture).PadLeft(3, '0') + date.ToString("/MM/dd", culture);

            return result;
        }

        /// <summary>
        /// 民國轉西元：091/01/01 => 2004-12-20 00:00:00.000
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static DateTime ToBCDate1(string date)
        {
            DateTime result = DateTime.MinValue;

            try
            {
                CultureInfo culture = new CultureInfo("zh-TW");
                culture.DateTimeFormat.Calendar = new System.Globalization.TaiwanCalendar();
                result = DateTime.Parse(date, culture);
            }
            catch (Exception ex)
            {
                result = DateTime.MinValue;
            }

            return result;
        }

        /// <summary>
        /// 民國轉西元：112年1月7日(六) => 2004-12-20 00:00:00.000
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static DateTime ToBCDate2(string date)
        {
            DateTime result = DateTime.MinValue;

            try
            {
                date = date.Split('日')[0].Replace('年', '/').Replace('月', '/');
                result = ToBCDate1(date);
            }
            catch (Exception ex)
            {
                result = DateTime.MinValue;
            }

            return result;
        }

        /// <summary>
        /// 民國轉西元：2023-05-08 (一) => 2004-12-20 00:00:00.000
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static DateTime ToBCDate3(string date)
        {
            DateTime result = DateTime.MinValue;

            try
            {
                date = date.Split(' ')[0];
                result = DateTime.Parse(date);
            }
            catch (Exception ex)
            {
                result = DateTime.MinValue;
            }

            return result;
        }

        /// <summary>
        /// 01:02
        /// </summary>
        /// <param name="time">01:02:00</param>
        /// <returns></returns>
        public static string ToTime1(string str)
        {
            string result = "";

            TimeSpan time;
            if (TimeSpan.TryParse(str, out time))
            {
                result = time.ToString(@"hh\:mm");
            }
            else
            {
                result = str;
            }

            return result;
        }
    }
}