using System;
using System.Collections.Generic;


namespace EFCoreDemo
{
    public class Leader
    {
        /// <summary>
        /// 自增ID
        /// </summary>
        public int Id { get; set; }

        public int WorkUnitId { get; set; }
        /// <summary>
        /// 序号
        /// </summary>
        public int OrderId { get; set; }
        /// <summary>
        /// 领导干部姓名
        /// </summary>
        public string FullName { get; set; }
        /// <summary>
        /// 领导干部身份证号
        /// </summary>
        public string CardNo { get; set; }
        /// <summary>
        /// 领导干部职务
        /// </summary>
        public string Post { get; set; }
        /// <summary>
        /// 领导干部职级
        /// </summary>
        public string Rank { get; set; }
        /// <summary>
        /// 是否存在经商办企业行为
        /// </summary>
        public bool IsExist { get; set; }
        /// <summary>
        /// 是否属于应予规范的情形
        /// </summary>
        public bool IsRegulated { get; set; }
        /// <summary>
        /// 退出方式
        /// </summary>
        public string ExitModel { get; set; }
        /// <summary>
        /// 是否比对
        /// </summary>
        public bool IsCompare { get; set; }
        /// <summary>
        /// 是否一致
        /// </summary>
        public bool IsMatch { get; set; }

        public DateTime CreateTime { get; set; }

        public WorkUnit WorkUnit { get; set; }

        public List<FamilyMember> FamilyMembers {get;set;}

    }
}
