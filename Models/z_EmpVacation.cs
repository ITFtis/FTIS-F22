using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    [Table("z_EmpVacation")]
    public class z_EmpVacation
    {
        //員工編號, 到職日期, 特休起算日期, 給假類型, 給假年度, 給假使用開始日期, 給假使用結束日期, 給假單位, 給假數量, 已用數量

        [Key]
        public string 員工編號 { get; set; }

        public DateTime? 到職日期 { get; set; }

        public DateTime? 特休起算日期 { get; set; }

        public string 給假類型 { get; set; }

        public int? 給假年度 { get; set; }

        public DateTime? 給假使用開始日期 { get; set; }

        public DateTime? 給假使用結束日期 { get; set; }

        public string 給假單位 { get; set; }

        public double? 給假數量 { get; set; }

        public double? 已用數量 { get; set; }

    }
}