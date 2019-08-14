using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace EFCoreDemo
{
    public class FamilyMember
    {
        /// <summary>
        /// 自增ID
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 对应的领导ID
        /// </summary>
        public int LeaderId { get; set; }

        /// <summary>
        /// 关系
        /// </summary>
        public string Relation { get; set; }
        /// <summary>
        /// 姓名
        /// </summary>
        public string FullName { get; set; }
        /// <summary>
        /// 身份证号
        /// </summary>
        public string CardNo { get; set; }
        /// <summary>
        /// 工作单位
        /// </summary>
        public string WorkUnit { get; set; }
        /// <summary>
        /// 身份证号是否通过核验
        /// </summary>
        public bool IsValid { get; set; }
        public DateTime CreateTime { get; set; }

        public Leader Leader { get; set; }

        public List<ReportBusiness> ReportBusinesses {get;set;}
    }
}
