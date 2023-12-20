using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    [Table("z_EmpOutTime")]
    public class z_EmpOutTime
    {
        /// <summary>
        /// 請假
        /// </summary>

        [Key]
        [Column(Order = 1)]
        public string 員工編號 { get; set; }
        [Key]
        [Column(Order = 2)]
        public string 單據編號 { get; set; }
        public string 公司代碼 { get; set; }
        public string 假別編號 { get; set; }
        public string 假別名稱 { get; set; }
        public DateTime? 開始日期 { get; set; }
        public string 開始時間 { get; set; }
        public DateTime? 結束日期 { get; set; }
        public string 結束時間 { get; set; }
        public Double 請假時數合計 { get; set; }
        public DateTime 申請日期 { get; set; }
        public string 請假原因 { get; set; }
    }
}