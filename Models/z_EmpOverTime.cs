using Dou.Misc.Attr;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    [Table("z_EmpOverTime")]
    public class z_EmpOverTime
    {
        /// <summary>
        /// 加班
        /// </summary>

        [Key]
        [Column(Order = 1)]
        public string 員工編號 { get; set; }
        [Key]
        [Column(Order = 2)]
        public string 加班單號 { get; set; }
        public string 專案編號 { get; set; }
        public string 處理方式 { get; set; }
        public DateTime 加班日期 { get; set; }
        public Double 加班時數A { get; set; }
        public Double 加班時數B { get; set; }
        public Double 加班總時數 { get; set; }
        public string 加班原因 { get; set; }
    }
}