using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    [Table("z_Comp")]
    public class z_Comp
    {
        /// <summary>
        /// 公司代碼
        /// </summary>

        [Key]
        public string 代碼 { get; set; }
        public string 名稱 { get; set; }
        public int 排序 { get; set; }
    }
}