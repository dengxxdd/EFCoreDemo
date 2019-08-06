using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFCoreDemo
{
    public class ReportBusiness
    {
        public int Id { get; set; }
        /// <summary>
        /// 经商办企业类型
        /// </summary>
        public string BusinessType{get;set;}
        /// <summary>
        /// 企业名称
        /// </summary>
        public string BusinessName { get; set; }
        /// <summary>
        /// 注册地
        /// </summary>
        public string RegPlace { get; set; }
        /// <summary>
        /// 经营地
        /// </summary>
        public string BusinessPlace { get; set; }
        /// <summary>
        /// 在企业所任职务
        /// </summary>
        public string BusinessPost { get; set; }
        /// <summary>
        /// 注册资本
        /// </summary>
        public decimal RegCapital { get; set; }
        /// <summary>
        /// 实缴资本
        /// </summary>
        public decimal PaidinCapital { get; set; }

        /// <summary>
        /// 成立日期
        /// </summary>
        public DateTime EstDate { get; set; }
        /// <summary>
        /// 近三年排名前5名销售商品或服务名称
        /// </summary>

        public string Nearly3Years1 { get; set; }
        /// <summary>
        /// 近三年排名前5名销售商品或服务金额
        /// </summary>
        public string Nearly3Years2 { get; set; }
        /// <summary>
        /// 近三年排名前5名销售商品或服务地域（到地级市）
        /// </summary>
        public string Nearly3Years3 { get; set; }
        /// <summary>
        /// 备注
        /// </summary>
        public string Remarks { get; set; }

        public FamilyMember FamilyMember { get; set; }
        public DateTime CreateTime { get; set; }

    }
}
