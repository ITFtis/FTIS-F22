using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace HRM_F22.Models
{
    [Table("z_Dep")]
    public class z_Dep
    {
        /// <summary>
        /// 部門代碼
        /// </summary>

        [Key]
        public string 編號 { get; set; }
        public string 名稱 { get; set; }
        public string 頁籤名稱 { get; set; }
        public int? 部門頁籤排序 { get; set; }
        /// <summary>
        /// 部門排序
        /// </summary>
        public int? 部門排序 { get; set; }
    }
}