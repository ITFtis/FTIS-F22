using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    [Table("CodeVacation")]
    public class CodeVacation
    {
        /// <summary>
        /// 假別編號
        /// </summary>

        [Key]
        public string 編號 { get; set; }
        public string 名稱 { get; set; }
        public string 類別名稱 { get; set; }
    }
}