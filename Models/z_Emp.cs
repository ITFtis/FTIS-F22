using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    [Table("z_Emp")]
    public class z_Emp
    {
        /// <summary>
        /// 員工
        /// </summary>
        
        [Key]
        public string 員工編號 { get; set; }
        public string 員工姓名 { get; set; }
        public string 部門編號 { get; set; }
        public string 部門名稱 { get; set; }
        public DateTime 到職日期 { get; set; }
        public DateTime 特休起算日期 { get; set; }
        public DateTime? 離職日期 { get; set; }
    }
}