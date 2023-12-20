using Dou.Controllers;
using HRM_F22.Models;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HRM_F22.Controllers
{
    public class TestBrianController : AGenericModelController<ReportData.SalaryOvertimeDetail>
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private string _errorDetail = "";

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

            try
            {
                string folder = FileHelper.GetFileFolder(Code.TempUploadFile.T8薪資結算);

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                //換薪加班
                if (!DoOverTimeMoney(folder, ref downloads))
                {
                    string errorMessage = "Excel產出錯誤，換薪加班" + "\n原因：" + _errorDetail; ;
                    logger.Info(errorMessage);
                    return Json(new { result = false, errorMessage = errorMessage }, JsonRequestBehavior.AllowGet);
                }

                //專案加班
                if (!DoOverTimeProject(folder, ref downloads))
                {
                    string errorMessage = "Excel產出錯誤，專案加班" + "\n原因：" + _errorDetail; ;
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

        public bool DoOverTimeMoney(string folder, ref List<string> downloads)
        {
            bool result = false;

            try
            {
                using (var db = new FtisT8DataSupplyModelContext())
                {
                    //部門名稱取代部門代碼顯示
                    var datas = db.Database.SqlQuery<ReportData.SalaryOvertimeDetail>(
                                    @"  Select 人員編號, 人員姓名, 所屬專案, 專案名稱, 
                                               b.名稱 AS 所屬部門代碼, 
                                               加班處理方案, A1, B1, A2, B2, B3, B4 
                                        From [dbo].[全月加班明細表] a
                                        Left Join 部門代碼 b On a.所屬部門代碼 = b.代碼                                        
                                        Where [加班處理方案] = '給加班費'
                                      ").ToArray();

                    if (datas.Count() == 0)
                    {
                        _errorDetail = "SQL查詢執行0筆";
                        return result;
                    }

                    //  Select[人員編號],[人員姓名],[所屬部門代碼],
                    //Sum(Convert(float, A1) + Convert(float, A2)) AS[換薪1],
                    //Sum(Convert(float, B1) + Convert(float, B2) + Convert(float, B3) + Convert(float, B4)) AS[換薪2]
                    //  From[全月加班明細表]
                    //  Where 1 = 1
                    //  And[加班處理方案] = '給加班費'
                    //  Group By[人員編號],[人員姓名],[所屬部門代碼]

                    var sheet = datas.GroupBy(a => new { 所屬部門代碼 = a.所屬部門代碼, 人員編號 = a.人員編號, 人員姓名 = a.人員姓名 })
                                .Select(a => new
                                {
                                    所屬部門代碼 = a.Key.所屬部門代碼,
                                    人員編號 = a.Key.人員編號,
                                    人員姓名 = a.Key.人員姓名,
                                    換薪1 = a.Sum(x => double.Parse(x.A1) + double.Parse(x.A2)),
                                    換薪2 = a.Sum(x => double.Parse(x.B1) + double.Parse(x.B2) + double.Parse(x.B3) + double.Parse(x.B4)),
                                    合計 = a.Sum(x => double.Parse(x.A1) + double.Parse(x.A2))
                                           + a.Sum(x => double.Parse(x.B1) + double.Parse(x.B2) + double.Parse(x.B3) + double.Parse(x.B4)),
                                    //SheetName = "換薪加班"  //設定sheet
                                });

                    sheet = sheet.OrderBy(a => a.所屬部門代碼).ThenBy(a => a.人員編號);
                    List<dynamic> list = new List<dynamic>();
                    //list.Add(sheet);
                    foreach (var row in sheet)
                    {
                        dynamic f = new ExpandoObject();
                        f.所屬部門代碼 = row.所屬部門代碼;
                        f.人員編號 = row.人員編號;
                        f.人員姓名 = row.人員姓名;
                        f.換薪1 = row.換薪1;
                        f.換薪2 = row.換薪2;
                        f.合計 = row.合計;
                        f.SheetName = "換薪加班";
                        list.Add(f);
                    }

                    string fileName = HRM_F22.ExcelHelper.GenerateExcelByLinq("換薪加班", list, folder);
                    string path = folder + fileName;
                    downloads.Add(HRM_F22.Cm.PhysicalToUrl(path));
                }
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

        public bool DoOverTimeProject(string folder, ref List<string> downloads)
        {
            bool result = false;

            try
            {
                using (var db = new FtisT8DataSupplyModelContext())
                {
                    //部門名稱取代部門代碼顯示
                    var datas = db.Database.SqlQuery<ReportData.SalaryOvertimeDetail>(
                                    @"  Select 人員編號, 人員姓名, 所屬專案, 專案名稱, 
                                               b.名稱 AS 所屬部門代碼, 
                                               加班處理方案, A1, B1, A2, B2, B3, B4 
                                        From [dbo].[全月加班明細表] a
                                        Left Join 部門代碼 b On a.所屬部門代碼 = b.代碼                                        
                                        Where [加班處理方案] = '給加班費'
                                      ").ToArray();

                    //var distDatas = datas.Select(a => new { a.人員編號, a.人員姓名 }).Distinct();
                    List<dynamic> list = new List<dynamic>();

                    //sheet區分：所屬部門代碼
                    var sheets = datas.GroupBy(x => x.所屬部門代碼).OrderBy(a => a.Key);
                    foreach (var sheet in sheets)
                    {
                        var data = datas.Where(a => a.所屬部門代碼 == sheet.First().所屬部門代碼).OrderBy(a => a.所屬專案);
                        //行列互換
                        var pivot = data.ToPivotArray(
                                       item => item.所屬專案,
                                       item => item.人員編號,
                                       v => v.Any() ? v.Sum(x => double.Parse(x.A1) + double.Parse(x.A2) + double.Parse(x.B1) + double.Parse(x.B2) + double.Parse(x.B3) + double.Parse(x.B4)) : 0);

                        foreach (var row in pivot.OrderBy(x => x.人員編號))
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
                            row.人員姓名 = datas.Where(a => a.人員編號 == row.人員編號).First().人員姓名;

                            //////設定sheet
                            ////row.SheetName = sheet.First().所屬部門代碼;
                            ////list.Add(row);

                            ////排序
                            dynamic f = new ExpandoObject();
                            f.人員編號 = row.人員編號;
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
                            f.SheetName = sheet.First().所屬部門代碼;
                            list.Add(f);
                        }
                    }

                    string fileName = HRM_F22.ExcelHelper.GenerateExcelByLinq("專案加班", list, folder);
                    string path = folder + fileName;
                    downloads.Add(HRM_F22.Cm.PhysicalToUrl(path));
                }
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