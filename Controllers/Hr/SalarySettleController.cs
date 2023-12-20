using Dou.Controllers;
using DouHelper;
using HRM_F22.Models;
using Microsoft.Ajax.Utilities;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

namespace HRM_F22.Controllers.Hr
{
    [Dou.Misc.Attr.MenuDef(Id = "SalarySettle", Name = "薪資結算用報表", MenuPath = "人資專區", Action = "Index", Index = 1, Func = Dou.Misc.Attr.FuncEnum.ALL, AllowAnonymous = false)]
    public class SalarySettleController : AGenericModelController<ReportData.SalaryOvertimeDetail>
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private string _errorDetail = "";

        Dou.Models.DB.IModelEntity<z_Emp> _emp = null;
        Dou.Models.DB.IModelEntity<z_EmpOverTime> _empOverTime = null;
        Dou.Models.DB.IModelEntity<z_Dep> _dep = null;
        Dou.Models.DB.IModelEntity<z_Comp> _comps = null;        
        Dou.Models.DB.IModelEntity<z_EmpOutTime> _empOutTime = null;
        Dou.Models.DB.IModelEntity<CodeVacation> _codeVacation = null;
        Dou.Models.DB.IModelEntity<z_EmpVacation> _empVacation = null;

        // GET: HrSalaryV2
        public ActionResult Index()
        {
            return View();
        }

        internal static System.Data.Entity.DbContext _dbContext = new FtisT8DataSupplyModelContext();
        protected override Dou.Models.DB.IModelEntity<ReportData.SalaryOvertimeDetail> GetModelEntity()
        {
            return new Dou.Models.DB.ModelEntity<ReportData.SalaryOvertimeDetail>(_dbContext);
        }

        public ActionResult DownloadFiles()
        {
            List<string> downloads = new List<string>();

            //查詢期間不可跨越，不然(扣事病假統計表_財務室)計算會有問題
            DateTime searchSDate = DateTime.Parse("2023/07/01");
            DateTime searchEDate = DateTime.Parse("2023/07/31");

            System.Data.Entity.DbContext db = new FtisT8DataSupplyModelContext();
            _emp = new Dou.Models.DB.ModelEntity<z_Emp>(db);
            _empOverTime = new Dou.Models.DB.ModelEntity<z_EmpOverTime>(db);
            _dep = new Dou.Models.DB.ModelEntity<z_Dep>(db);
            _comps = new Dou.Models.DB.ModelEntity<z_Comp>(db);
            _empOverTime = new Dou.Models.DB.ModelEntity<z_EmpOverTime>(db);
            _empOutTime = new Dou.Models.DB.ModelEntity<z_EmpOutTime>(db);
            _codeVacation = new Dou.Models.DB.ModelEntity<CodeVacation>(db);
            _empVacation = new Dou.Models.DB.ModelEntity<z_EmpVacation>(db);

            try
            {
                string folder = FileHelper.GetFileFolder(Code.TempUploadFile.T8薪資結算);

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                //報表資料回傳
                List<dynamic> dyListOverTimeMoney = null;
                List<dynamic> dyListStuffTime = null;

                //換薪加班
                if (!DoOverTimeMoney(folder, ref downloads, ref dyListOverTimeMoney))
                {
                    string errorMessage = "換薪加班" + "\n原因：" + _errorDetail; ;
                    logger.Info(errorMessage);
                    return Json(new { result = false, errorMessage = errorMessage }, JsonRequestBehavior.AllowGet);
                }

                //專案加班
                if (!DoOverTimeProject(folder, ref downloads))
                {
                    string errorMessage = "專案加班" + "\n原因：" + _errorDetail; ;
                    logger.Info(errorMessage);
                    return Json(new { result = false, errorMessage = errorMessage }, JsonRequestBehavior.AllowGet);
                }

                //扣事病假統計表_財務室
                if (!DoStuffTime(folder, ref downloads, ref dyListStuffTime, searchSDate))
                {
                    string errorMessage = "扣事病假統計表_財務室" + "\n原因：" + _errorDetail; ;
                    logger.Info(errorMessage);
                    return Json(new { result = false, errorMessage = errorMessage }, JsonRequestBehavior.AllowGet);
                }

                //同仁離職結算
                if (!DoQuit(folder, ref downloads, searchSDate))
                {
                    string errorMessage = "同仁離職結算" + "\n原因：" + _errorDetail; ;
                    logger.Info(errorMessage);
                    return Json(new { result = false, errorMessage = errorMessage }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                logger.Info("系統執行錯誤：" + ex.Message + " " + ex.InnerException);
                logger.Info(ex.StackTrace);

                return Json(new { result = false, errorMessage = "系統執行錯誤" });
            }

            return Json(new { result = true, url = downloads }, JsonRequestBehavior.AllowGet);
        }

        //換薪加班
        public bool DoOverTimeMoney(string folder, ref List<string> downloads, ref List<dynamic> dyListOverTimeMoney)
        {
            bool result = false;

            try
            {
                var datas = _emp.GetAll().Where(a => a.離職日期 == null)
                    .Join(_dep.GetAll(), a => a.部門編號, b => b.編號, (o, c) => new
                    {
                        o.員工編號,
                        o.員工姓名,
                        部門名稱 = c.名稱,
                        部門頁籤排序 = c.部門頁籤排序 != null ? c.部門頁籤排序 : 999,
                        部門排序 = c.部門排序 != null ? c.部門排序 : 999
                    })
                    .GroupJoin(_empOverTime.GetAll(), a => a.員工編號, b => b.員工編號, (o, c) => new { o.員工編號, o.部門名稱, o.員工姓名, o.部門頁籤排序, o.部門排序, c })
                    .SelectMany(o => o.c.DefaultIfEmpty(), (o, c) =>
                    new
                    {
                        公司代碼 = o.員工編號.Substring(0, 1),
                        o.員工編號,
                        o.部門名稱,
                        o.部門頁籤排序,
                        o.部門排序,
                        o.員工姓名,
                        加班單號 = c.加班單號 ?? null,
                        專案編號 = c.專案編號 ?? null,
                        處理方式 = c.處理方式 ?? null,
                        加班日期 = c.加班日期,
                        加班時數A = c.加班時數A,
                        加班時數B = c.加班時數B,
                        加班原因 = c.加班原因 ?? null
                    }).Where(c => c.加班單號 != null)
                    .Where(p => p.處理方式 == "給加班費")
                    .ToList();

                ////if (datas.Count() == 0)
                ////{
                ////    _errorDetail = "加班資料0筆";
                ////    return result;
                ////}

                var gpDatas = datas.GroupBy(a => new { a.公司代碼, a.部門名稱, a.部門頁籤排序, a.部門排序, a.員工編號, a.員工姓名 })
                            .Select(a => new
                            {
                                公司代碼 = a.Key.公司代碼,
                                部門名稱 = a.Key.部門名稱,
                                部門頁籤排序 = a.Key.部門頁籤排序,
                                部門排序 = a.Key.部門排序,
                                員工編號 = a.Key.員工編號,
                                員工姓名 = a.Key.員工姓名,
                                換薪1 = a.Sum(x => x.加班時數A),
                                換薪2 = a.Sum(x => x.加班時數B),
                                合計 = a.Sum(x => x.加班時數A) + a.Sum(x => x.加班時數B)
                                //SheetName = "換薪加班"  //設定sheet
                            })
                            .OrderBy(a => a.部門名稱).ThenBy(a => a.員工編號);

                List<dynamic> list = new List<dynamic>();

                //sheet區分：公司名稱                
                foreach (var sheep in _comps.GetAll().OrderBy(a => a.排序))
                {
                    var data = gpDatas.Where(a => a.公司代碼 == sheep.代碼)
                                    .OrderBy(a => a.部門頁籤排序).ThenBy(b => b.部門排序);

                    if (data.Count() == 0)
                    {
                        //當下查詢年月 沒資料
                        dynamic f = new ExpandoObject();
                        f.SheetName = sheep.名稱;
                        list.Add(f);
                        continue;
                    }
                    else
                    {
                        foreach (var row in data)
                        {
                            dynamic f = new ExpandoObject();
                            f.所屬部門代碼 = row.部門名稱;
                            f.人員編號 = row.員工編號;
                            f.人員姓名 = row.員工姓名;
                            f.換薪1 = row.換薪1;
                            f.換薪2 = row.換薪2;
                            f.合計 = row.合計;

                            //設定sheet
                            f.SheetName = sheep.名稱;
                            list.Add(f);
                        }
                    }
                }

                //報表資料回傳
                dyListOverTimeMoney = list;

                string fileName = HRM_F22.ExcelHelper.GenerateExcelByLinq("換薪加班", list, folder);
                string path = folder + fileName;
                downloads.Add(HRM_F22.Cm.PhysicalToUrl(path));
            }
            catch (Exception ex)
            {
                logger.Info(ex.Message + " " + ex.InnerException);
                logger.Info(ex.StackTrace);

                return result;
            }

            result = true;

            return result;
        }

        //專案加班
        public bool DoOverTimeProject(string folder, ref List<string> downloads)
        {
            bool result = false;

            try
            {
                var datas = _emp.GetAll().Where(a => a.離職日期 == null)
                    .Join(_dep.GetAll(), a => a.部門編號, b => b.編號, (o, c) => new
                    {
                        o.員工編號,
                        o.員工姓名,
                        部門名稱 = c.名稱,
                        頁籤名稱 = !string.IsNullOrEmpty(c.頁籤名稱) ? c.頁籤名稱 : c.名稱,
                        部門頁籤排序 = c.部門頁籤排序 != null ? c.部門頁籤排序 : 999,
                        部門排序 = c.部門排序 != null ? c.部門排序 : 999
                    })
                    .GroupJoin(_empOverTime.GetAll(), a => a.員工編號, b => b.員工編號, (o, c) => new { o.員工編號, o.部門名稱, o.員工姓名, o.頁籤名稱, o.部門頁籤排序, o.部門排序, c })
                    .SelectMany(o => o.c.DefaultIfEmpty(), (o, c) =>
                    new
                    {
                        公司代碼 = o.員工編號.Substring(0, 1),
                        o.員工編號,
                        o.部門名稱,
                        o.頁籤名稱,
                        o.部門頁籤排序,
                        o.部門排序,
                        o.員工姓名,
                        加班單號 = c.加班單號 ?? null,
                        專案編號 = c.專案編號 ?? null,
                        處理方式 = c.處理方式 ?? null,
                        加班日期 = c.加班日期,
                        加班時數A = c.加班時數A,
                        加班時數B = c.加班時數B,
                        加班原因 = c.加班原因 ?? null
                    }).Where(c => c.加班單號 != null)
                    .Where(p => p.處理方式 == "給加班費")
                    .ToList();

                ////if (datas.Count() == 0)
                ////{
                ////    _errorDetail = "加班資料0筆";
                ////    return result;
                ////}

                List<dynamic> list = new List<dynamic>();

                //sheet區分：(部門名稱-沒頁籤的資料)+(頁籤名稱)
                var s1 = datas.Select(a => new { a.頁籤名稱, a.部門頁籤排序 }).Distinct().ToList();
                var s2 = _dep.GetAll().Where(a => !string.IsNullOrEmpty(a.頁籤名稱))
                    .Select(a => new { a.頁籤名稱, a.部門頁籤排序 }).Distinct().ToList();
                var tabs = s1.Union(s2);
                foreach (var sheet in tabs.OrderBy(a => a.部門頁籤排序))
                {
                    var data = datas.Where(a => a.頁籤名稱 == sheet.頁籤名稱).OrderBy(a => a.專案編號);

                    if (data.Count() == 0)
                    {
                        //當下查詢年月 沒資料
                        dynamic f = new ExpandoObject();
                        f.SheetName = sheet.頁籤名稱;
                        list.Add(f);
                        continue;
                    }

                    //行列互換
                    var pivot = data.ToPivotArray(
                                   item => item.專案編號,
                                   item => item.員工編號,
                                   v => v.Any() ? v.Sum(x => x.加班時數A + x.加班時數B) : 0);

                    foreach (var row in pivot
                        .OrderBy(row => datas.Where(a => a.員工編號 == row.員工編號).First().部門排序)
                        .ThenBy(a => a.員工編號))
                    {
                        double total = 0;

                        foreach (var item in (dynamic)row)
                        {
                            string val = item.Value.ToString();

                            double value = 0;
                            double.TryParse(val, out value);
                            total += value;
                        }

                        row.合計 = total;
                        row.部門名稱 = datas.Where(a => a.員工編號 == row.員工編號).First().部門名稱;
                        row.人員姓名 = datas.Where(a => a.員工編號 == row.員工編號).First().員工姓名;

                        ////排序
                        dynamic f = new ExpandoObject();
                        f.部門名稱 = row.部門名稱;
                        f.員工編號 = row.員工編號;
                        f.人員姓名 = row.人員姓名;
                        f.合計 = row.合計;
                        //(動態數字)                            
                        foreach (var item in (dynamic)row)
                        {
                            string key = item.Key.ToString();
                            string val = item.Value.ToString();

                            if (((IDictionary<string, object>)f).Keys.Contains(key))
                                continue;

                            double value = 0;
                            if (double.TryParse(val, out value))
                            {
                                KeyValuePair<string, object> r = new KeyValuePair<string, object>(key, value);
                                ((IDictionary<string, object>)f).Add(r);
                            }
                        }

                        //設定sheet
                        f.SheetName = sheet.頁籤名稱;
                        list.Add(f);
                    }
                }

                string fileName = HRM_F22.ExcelHelper.GenerateExcelByLinq("專案加班", list, folder);
                string path = folder + fileName;
                downloads.Add(HRM_F22.Cm.PhysicalToUrl(path));
            }
            catch (Exception ex)
            {
                logger.Info(ex.Message + " " + ex.InnerException);
                logger.Info(ex.StackTrace);

                return result;
            }

            result = true;

            return result;
        }

        //扣事病假統計表_財務室
        public bool DoStuffTime(string folder, ref List<string> downloads, ref List<dynamic> dyListStuffTime, DateTime searchSDate)
        {
            bool result = false;

            try
            {
                var empOut = _empOutTime.GetAll().Join(_codeVacation.GetAll(), a => a.假別編號, b => b.編號, (s, t) =>
                new
                {
                    員工編號 = s.員工編號 ?? null,
                    單據編號 = s.單據編號 ?? null,
                    假別名稱 = s.假別名稱 ?? null,
                    開始日期 = s.開始日期 != null ? s.開始日期 : DateTime.MinValue,
                    開始時間 = s.開始時間 ?? null,
                    結束日期 = s.結束日期 != null ? s.結束日期 : DateTime.MaxValue,
                    結束時間 = s.結束時間 ?? null,
                    請假時數合計 = (Double)s.請假時數合計,
                    請假原因 = s.請假原因 ?? null,
                    請假類別名稱 = t.類別名稱
                });

                //data2s(當年、前1年、前2年)  累計前2年度的事病假資料
                var data2s = _emp.GetAll().Where(a => a.離職日期 == null)
                    .Join(_comps.GetAll(), a => a.員工編號.Substring(0, 1), b => b.代碼, (s, t) => new { 公司名稱 = t.名稱, 公司排序 = t.排序, s.員工編號, s.部門名稱, s.員工姓名, s.到職日期 })
                    .GroupJoin(empOut, a => a.員工編號, b => b.員工編號, (o, c) => new { o.公司名稱, o.公司排序, o.員工編號, o.部門名稱, o.員工姓名, o.到職日期, c })
                    .SelectMany(o => o.c.DefaultIfEmpty(), (o, c) =>
                    new
                    {
                        序號 = Guid.NewGuid(),
                        o.公司名稱,
                        o.公司排序,
                        o.員工編號,
                        o.部門名稱,
                        o.員工姓名,
                        o.到職日期,
                        病假重算起日期 = (DateTime)DbFunctions.CreateDateTime(searchSDate.Year - 1, o.到職日期.Month, o.到職日期.Day, 0, 0, 0.0),
                        病假重算迄日期 = (DateTime)DbFunctions.AddDays(DbFunctions.CreateDateTime(searchSDate.Year, o.到職日期.Month, o.到職日期.Day, 0, 0, 0.0), -1),
                        c.單據編號,
                        c.請假類別名稱,
                        c.開始日期,
                        c.開始時間,
                        c.結束日期,
                        c.結束時間,
                        c.請假時數合計,
                        c.請假原因
                    }).Where(c => c.單據編號 != null)
                    .Where(a => a.請假類別名稱 == "事假" || a.請假類別名稱 == "病假")
                    .Where(a => searchSDate.Year - a.開始日期.Value.Year <= 2)//前2年請假資料
                    .ToList();

                //////if (data2s.Count() == 0)
                //////{
                //////    _errorDetail = "事病假資料0筆";
                //////    return result;
                //////}

                List<dynamic> list = new List<dynamic>();
                List<dynamic> outDetails = new List<dynamic>();

                //sheet區分：公司名稱
                var sheets = data2s.GroupBy(x => new { x.公司名稱, x.公司排序 })
                    .OrderBy(a => a.Key.公司排序)
                    .Select(a => new { a.Key.公司名稱 });

                foreach (var sheet in sheets)
                {
                    //data2s(當年、前1年、前2年)  累計前2年度的事病假資料
                    //當下查詢年月
                    var data = data2s.Where(a => a.公司名稱 == sheet.公司名稱)
                                .Where(a => searchSDate.Year == a.開始日期.Value.Year && searchSDate.Month == a.開始日期.Value.Month);

                    if (data.Count() == 0)
                    {
                        //當下查詢年月 沒資料
                        dynamic f = new ExpandoObject();
                        f.SheetName = sheet.公司名稱;
                        list.Add(f);
                        continue;
                    }

                    //(1)員工事病假應扣時數計算(已扣+未扣)
                    //a.(已扣統計)非本月累計病假時數
                    //條件：病假累計重算，病假開始日期：到職日(+1年/月/日)
                    Dictionary<string, double> dicDidSicks =
                        data2s.Where(a => !data.Any(b => a.序號 == b.序號))
                        .Where(a => a.請假類別名稱 == "病假")
                        .Where(a => a.開始日期 >= a.病假重算起日期 && a.開始日期 <= a.病假重算迄日期)
                        .GroupBy(a => a.員工編號)
                        .Select(a => new
                        {
                            員工編號 = a.Key,
                            病假已扣時數 = (double)a.Sum(b => b.請假時數合計)
                        }).ToDictionary(x => x.員工編號, x => x.病假已扣時數);

                    //(未扣細項)本月事病假時數
                    List<ReportData.NoDidSicksDetail> listNoDidSicksDetail = new List<ReportData.NoDidSicksDetail>();
                    //以天為單位，累計(未扣)病假時數
                    Dictionary<string, double> dicLog = new Dictionary<string, double>();
                    //以天為單位，病假重算名單
                    Dictionary<string, string> dicReset = new Dictionary<string, string>();

                    foreach (var v in data.OrderBy(a => a.員工編號).ThenBy(a => a.開始日期))
                    {
                        if (!dicDidSicks.ContainsKey(v.員工編號))
                            dicDidSicks.Add(v.員工編號, 0);

                        //符合重算病假日期(條件加1個：病假重算大於當月份查詢才要)
                        if (v.病假重算迄日期 >= searchSDate && v.開始日期 > v.病假重算迄日期)
                        {
                            if (!dicReset.ContainsKey(v.員工編號))
                            {
                                //尚未重算
                                dicDidSicks[v.員工編號] = 0;  //(已扣)病假時數                                
                                dicLog[v.員工編號] = 0;       //(未扣)紀錄累計病假時數

                                string memo = string.Format("{0} 病假重算日", DateFormat.ToDate1(v.病假重算迄日期.AddDays(1)));
                                dicReset.Add(v.員工編號, memo);
                            }
                        }

                        string sickType = "";  //A類人(96hr免扣錢)、B累人
                        if (sheet.公司名稱 == "產基會"
                            && v.到職日期.Year < 2023)
                            sickType = "A";
                        else
                            sickType = "B";

                        if (v.請假類別名稱 == "事假")
                        {
                            //事假 * 1
                            listNoDidSicksDetail.Add(new ReportData.NoDidSicksDetail()
                            {
                                員工編號 = v.員工編號,
                                病假類種人 = sickType,
                                當月病假時數96_240 = 0,
                                當月病假時數240以下 = 0,
                                當月病假時數240以上 = 0,
                                應扣時數 = v.請假時數合計 * 1
                            });
                        }
                        else if (v.請假類別名稱 == "病假")
                        {
                            //病假
                            if (!dicLog.ContainsKey(v.員工編號))
                                dicLog.Add(v.員工編號, 0);

                            double subHours = 0;
                            double 當月病假時數96_240 = 0;
                            double 當月病假時數240以下 = 0;
                            double 當月病假時數240以上 = 0;

                            //240hr以內 * 0.5...等規則
                            //1.原總和(10) + 8 =18(8hr * 0.5)
                            //2.原總和(240) + 8 =248(8hr * 0.5)
                            //3.原總和(235) + 8 =243(5hr * 0.5, 3hr * 1)
                            double a = dicDidSicks[v.員工編號];  //(已扣)病假時數                                
                            double b = dicLog[v.員工編號];       //(未扣)紀錄累計病假時數
                            double c = v.請假時數合計;
                            double sum = a + b + c;

                            if (sickType == "A")
                            {
                                //A類人
                                if (sum <= 96)
                                {
                                    //subHours = 0;
                                }
                                else if (sum <= 240)
                                {
                                    var num = sum - 96;
                                    if (num < v.請假時數合計)
                                    {
                                        當月病假時數96_240 = num * 0.5;
                                    }
                                    else
                                    {
                                        當月病假時數96_240 = v.請假時數合計 * 0.5;
                                    }
                                    subHours = 當月病假時數96_240;
                                }
                                else
                                {
                                    //230+8+8     (246 - 240) =>  扣6 * 1 + 2 * 0.5
                                    //230+8+8+5   (251 - 240) =>  扣5 * 1
                                    //230+8+8+8+2 (253 - 240) =>  扣2 * 1

                                    var num = v.請假時數合計 - (sum - 240);
                                    if (num > 0)
                                    {
                                        當月病假時數96_240 = num * 0.5;
                                        當月病假時數240以上 = (v.請假時數合計 - num) * 1;
                                        subHours = 當月病假時數240以上 + 當月病假時數96_240;
                                    }
                                    else
                                    {
                                        當月病假時數240以上 = v.請假時數合計 * 1;
                                        subHours = 當月病假時數240以上;
                                    }
                                }
                            }
                            else if (sickType == "B")
                            {
                                //B類人
                                if (sum <= 240)
                                {
                                    當月病假時數240以下 = v.請假時數合計 * 0.5;
                                    subHours = 當月病假時數240以下;
                                }
                                else
                                {
                                    //230+8+8     (246 - 240) =>  扣6 * 1 + 2 * 0.5
                                    //230+8+8+5   (251 - 240) =>  扣5 * 1
                                    //230+8+8+8+2 (253 - 240) =>  扣2 * 1

                                    var num = v.請假時數合計 - (sum - 240);
                                    if (num > 0)
                                    {
                                        當月病假時數240以下 = num * 0.5;
                                        當月病假時數240以上 = (v.請假時數合計 - num) * 1;
                                        subHours = 當月病假時數240以上 + 當月病假時數240以下;
                                    }
                                    else
                                    {
                                        當月病假時數240以上 = v.請假時數合計 * 1;
                                        subHours = 當月病假時數240以上;
                                    }
                                }
                            }

                            dicLog[v.員工編號] = dicLog[v.員工編號] + v.請假時數合計;
                            listNoDidSicksDetail.Add(new ReportData.NoDidSicksDetail()
                            {
                                員工編號 = v.員工編號,
                                病假類種人 = sickType,
                                當月病假時數96_240 = 當月病假時數96_240,
                                當月病假時數240以下 = 當月病假時數240以下,
                                當月病假時數240以上 = 當月病假時數240以上,
                                應扣時數 = subHours
                            });
                        }
                    }

                    //b.(未扣統計)本月事病假時數
                    var listNoDid = listNoDidSicksDetail.GroupBy(a => new { a.員工編號, a.病假類種人 })
                        .Select(g => new
                        {
                            員工編號 = g.Key.員工編號,
                            病假類種人 = g.Key.病假類種人,
                            當月病假時數96_240 = g.Sum(a => a.當月病假時數96_240),
                            當月病假時數240以下 = g.Sum(a => a.當月病假時數240以下),
                            當月病假時數240以上 = g.Sum(a => a.當月病假時數240以上),
                            應扣時數 = g.Sum(a => a.應扣時數),
                            備註 = dicReset.ContainsKey(g.Key.員工編號) ? dicReset[g.Key.員工編號] : ""
                        }); ;

                    //(2)產出Excel Row
                    var pivot = data.GroupBy(a => new { a.部門名稱, a.員工編號, a.員工姓名 })
                                .Select(a => new
                                {
                                    a.Key.部門名稱,
                                    a.Key.員工編號,
                                    a.Key.員工姓名,
                                    事假時數 = a.Sum(p => p.請假類別名稱 == "事假" ? p.請假時數合計 : 0),
                                    病假時數 = a.Sum(p => p.請假類別名稱 == "病假" ? p.請假時數合計 : 0)
                                })
                                .GroupJoin(listNoDid, a => a.員工編號, b => b.員工編號, (o, c) => new
                                {
                                    o.部門名稱,
                                    o.員工編號,
                                    o.員工姓名,
                                    o.事假時數,
                                    o.病假時數,
                                    c.First().病假類種人,
                                    c.First().當月病假時數96_240,
                                    c.First().當月病假時數240以下,
                                    c.First().當月病假時數240以上,
                                    c.First().應扣時數,
                                    備註 = c.First().備註
                                });

                    foreach (var row in pivot
                        .OrderBy(x => x.員工編號))
                    {
                        ////1.匯出(設定sheet)  請假統計
                        dynamic f = new ExpandoObject();
                        f.部門名稱 = row.部門名稱;
                        f.員工編號 = row.員工編號;
                        f.員工姓名 = row.員工姓名;

                        ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月事假", row.事假時數));
                        if (sheet.公司名稱 == "產基會")
                        {
                            f.累計病假超過96才登 = row.病假類種人 == "A" && dicDidSicks[f.員工編號] + row.病假時數 > 96 ? dicDidSicks[f.員工編號] + row.病假時數 : "";
                            f.超出免扣病假時數 = row.病假類種人 == "A" && dicDidSicks[f.員工編號] + row.病假時數 > 96 ? dicDidSicks[f.員工編號] + row.病假時數 - 96 : "";

                            f.已扣病假時數_含96 = dicDidSicks[f.員工編號]; //dicDidSicks[f.員工編號] - 96 > 0 ? dicDidSicks[f.員工編號] - 96 : "";
                            ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月病假", row.病假時數));
                            ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月A類應扣病假時數(96~240間)乘0.5", row.當月病假時數96_240));
                            ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月B類應扣病假時數(240以下)乘0.5", row.當月病假時數240以下));
                            ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月應扣病假時數(超過240)乘1", row.當月病假時數240以上));
                        }
                        else
                        {
                            ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月病假", row.病假時數));
                            ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月應扣病假時數(240以下)乘0.5", row.當月病假時數240以下));
                            ((IDictionary<string, object>)f).Add(new KeyValuePair<string, object>(searchSDate.Month.ToString() + "月應扣病假時數(超過240)乘1", row.當月病假時數240以上));
                        }

                        f.應扣時數 = row.應扣時數;
                        f.備註 = row.備註;

                        //設定sheet
                        f.SheetName = sheet.公司名稱;
                        list.Add(f);
                    }

                    int index = 0;
                    foreach (var row in data
                        .OrderBy(x => x.員工編號).ThenBy(b => b.開始日期))
                    {
                        ////2.匯出(設定sheet)  請假明細
                        //.									
                        index++;

                        dynamic f = new ExpandoObject();
                        f.公司 = row.公司名稱;
                        f.No = index;
                        f.職號 = row.員工編號;
                        f.姓名 = row.員工姓名;
                        f.組別 = row.部門名稱;
                        f.假別 = row.請假類別名稱;
                        f.開始日期 = row.開始日期;
                        f.開始時間 = row.開始時間;
                        f.結束日期 = row.結束日期;
                        f.結束時間 = row.結束時間;
                        f.時數 = row.請假時數合計;

                        //設定sheet
                        f.SheetName = "請假明細";
                        outDetails.Add(f);
                    }
                }

                list.AddRange(outDetails);

                //報表資料回傳
                dyListStuffTime = list;

                string fileName = HRM_F22.ExcelHelper.GenerateExcelByLinq("扣事病假統計表_財務室", list, folder);
                string path = folder + fileName;
                downloads.Add(HRM_F22.Cm.PhysicalToUrl(path));
            }
            catch (Exception ex)
            {
                logger.Info(ex.Message + " " + ex.InnerException);
                logger.Info(ex.StackTrace);

                return result;
            }

            result = true;

            return result;
        }

        //同仁離職結算
        public bool DoQuit(string folder, ref List<string> downloads, DateTime searchSDate)
        {
            bool result = false;

            try
            {
                var datas = _emp.GetAll().Where(a => a.離職日期 != null)
                    .GroupJoin(_empVacation.GetAll(), a => a.員工編號, b => b.員工編號, (o, c) => new { o.員工編號, o.部門名稱, o.員工姓名, o.到職日期, o.離職日期, c })
                    .SelectMany(o => o.c.DefaultIfEmpty(), (o, c) =>
                    new
                    {
                        公司代碼 = o.員工編號.Substring(0, 1),
                        o.員工編號,
                        o.員工姓名,
                        o.到職日期,
                        年資 = Math.Round(((double)DbFunctions.DiffMonths(o.到職日期, o.離職日期) / 12), 2),
                        o.離職日期
                    })
                    .Where(c => ((DateTime)c.離職日期).Year == searchSDate.Year
                             && ((DateTime)c.離職日期).Month == searchSDate.Month)

                    .ToList();

                List<dynamic> list = new List<dynamic>();

                //sheet區分：公司名稱                    
                foreach (var sheep in _comps.GetAll().OrderBy(a => a.排序))
                {
                    var data = datas.Where(a => a.公司代碼 == sheep.代碼);

                    if (data.Count() == 0)
                    {
                        //當下查詢年月 沒資料
                        dynamic f = new ExpandoObject();
                        f.SheetName = sheep.名稱;
                        list.Add(f);
                        continue;
                    }
                    else
                    {
                        ////var gpDatas = datas.GroupBy(a => new { a.公司代碼, a.部門名稱, a.部門頁籤排序, a.部門排序, a.員工編號, a.員工姓名 })
                        ////    .Select(a => new
                        ////    {
                        ////        公司代碼 = a.Key.公司代碼,
                        ////        部門名稱 = a.Key.部門名稱,
                        ////        部門頁籤排序 = a.Key.部門頁籤排序,
                        ////        部門排序 = a.Key.部門排序,
                        ////        員工編號 = a.Key.員工編號,
                        ////        員工姓名 = a.Key.員工姓名,
                        ////        換薪1 = a.Sum(x => x.加班時數A),
                        ////        換薪2 = a.Sum(x => x.加班時數B),
                        ////        合計 = a.Sum(x => x.加班時數A) + a.Sum(x => x.加班時數B)
                        ////        //SheetName = "換薪加班"  //設定sheet
                        ////    })
                        ////    .OrderBy(a => a.部門名稱).ThenBy(a => a.員工編號);

                        foreach (var row in data)
                        {
                            dynamic f = new ExpandoObject();
                            f.人員編號 = row.員工編號;
                            f.人員姓名 = row.員工姓名;
                            f.到職日期 = row.到職日期;
                            f.年資 = row.年資;

                            //設定sheet
                            f.SheetName = sheep.名稱;
                            list.Add(f);
                        }
                    }
                }

                string fileName = HRM_F22.ExcelHelper.GenerateExcelByLinq("同仁離職結算", list, folder);
                string path = folder + fileName;
                downloads.Add(HRM_F22.Cm.PhysicalToUrl(path));
            }
            catch (Exception ex)
            {
                logger.Info(ex.Message + " " + ex.InnerException);
                logger.Info(ex.StackTrace);

                return result;
            }

            result = true;

            return result;
        }
    }
}