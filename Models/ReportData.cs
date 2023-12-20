using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    public class ReportData
    {
        //人員編號, 人員姓名, 所屬專案, 專案名稱, 簽核主管, 主管姓名, 管理公司, 所屬部門代碼, 加班日期, 加班單號, 加班處理方案, 加班開始日期, 加班開始時間, 加班結束日期, 加班結束時間, A1, B1, A2, B2, B3, B4, 理論時數合計, 時數合計, 加班費時數合計, 補休時數合計, 已補休時數, 已折現時數, 剩餘可補休時數, 加班原因, 單位, 備註
        public class SalaryOvertimeDetail
        {
            public string 人員編號 { get; set; }
            public string 人員姓名 { get; set; }
            public string 所屬專案 { get; set; }
            public string 專案名稱 { get; set; }
            public string 所屬部門代碼 { get; set; }
            public string 加班處理方案 { get; set; }
            public string A1 { get; set; }
            public string B1 { get; set; }
            public string A2 { get; set; }
            public string B2 { get; set; }
            public string B3 { get; set; }
            public string B4 { get; set; }
        }

        public class sp_salary_overtimemoney
        {
            [Display(Name = "所屬部門代碼")]
            public string 所屬部門代碼 { get; set; }

            [Display(Name = "人員編號")]
            public string 人員編號 { get; set; }

            [Display(Name = "人員姓名")]
            public string 人員姓名 { get; set; }

            [Display(Name = "換薪1")]
            public double 換薪1 { get; set; }

            [Display(Name = "換薪2")]
            public double 換薪2 { get; set; }

            [Display(Name = "合計")]
            public double 合計 { get; set; }
        }

        //扣事病假統計表_財務室 => (未扣細項)本月事病假時數
        public class NoDidSicksDetail
        {
            public string 員工編號 { get; set; }
            public string 病假類種人 { get; set; }
            public double 當月病假時數96_240 { get; set; }
            public double 當月病假時數240以下 { get; set; }
            public double 當月病假時數240以上 { get; set; }
            public double 應扣時數 { get; set; }
        }
    }
}