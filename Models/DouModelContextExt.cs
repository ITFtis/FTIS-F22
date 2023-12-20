using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace HRM_F22.Models
{
    public class FtisT8PartModelContext : Dou.Models.ModelContextBase<User, Role>
    {
        /// <summary>
        /// T8部分獨立 ex:User
        /// </summary>
        public FtisT8PartModelContext() : base("name=FtisT8PartModelContext")
        {
            //Database.SetInitializer<FtisT8PartModelContext>(null);
        }
    }

    public class FtisT8DataSupplyModelContext : System.Data.Entity.DbContext
    {
        public FtisT8DataSupplyModelContext() : base("name=FtisT8DataSupplyModelContext")
        {
            Database.SetInitializer<FtisT8DataSupplyModelContext>(null);
        }

        public virtual DbSet<z_Emp> z_Emp { get; set; }
        public virtual DbSet<z_EmpOverTime> z_EmpOverTime { get; set; }
        public virtual DbSet<z_EmpOutTime> z_EmpOutTime { get; set; }
        public virtual DbSet<z_Dep> z_Dep { get; set; }
        public virtual DbSet<z_Comp> z_Comp { get; set; }
        public virtual DbSet<CodeVacation> CodeVacation { get; set; }

        public virtual DbSet<z_EmpVacation> z_EmpVacation { get; set; }
    }
}