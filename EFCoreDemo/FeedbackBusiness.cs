using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFCoreDemo
{
    public class FeedbackBusiness
    {
        public int Id { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string FullName { get; set; }
        /// <summary>
        /// 身份证号
        /// </summary>
        public string CardNo { get; set; }
        /// <summary>
        /// 企业名称
        /// </summary>
        public string BusinessName { get; set; }
        /// <summary>
        /// 经商办企业类型
        /// </summary>
        public string BusinessType { get; set; }
        /// <summary>
        /// 成立日期
        /// </summary>
        public DateTime EstDate { get; set; }
        /// <summary>
        /// 法定代表人
        /// </summary>
        public string LegalRep { get; set; }
        /// <summary>
        /// 住所
        /// </summary>
        public string Residence { get; set; }
        /// <summary>
        /// 登记状态
        /// </summary>
        public string State { get; set; }
        /// <summary>
        /// 登记状态
        /// </summary>
        public DateTime LogoffDate { get; set; }
        /// <summary>
        /// 吊销时间
        /// </summary>
        public DateTime RevokeDate { get; set; }
        /// <summary>
        /// 是否异常
        /// </summary>
        public bool IsAbnormal { get; set; }
        /// <summary>
        /// 是否严重违法
        /// </summary>
        public bool IsOutrage { get; set; }
        /// <summary>
        /// 统一信用代码
        /// </summary>
        public string CreditCode{get;set;}
        /// <summary>
        /// 注册资本
        /// </summary>
        public decimal RegisteredCapital { get; set; }
        /// <summary>
        /// 币种
        /// </summary>
        public string Currency { get; set; }
        /// <summary>
        /// 认缴出资额
        /// </summary>
        public decimal Subscribe { get; set; }
        /// <summary>
        /// 出资比例
        /// </summary>
        public decimal Proportion { get; set; }
        /// <summary>
        /// 出资时间
        /// </summary>
        public DateTime PutupDate { get; set; }
        /// <summary>
        /// 在企业所任职务
        /// </summary>
        public string BusinessPost { get; set; }
        /// <summary>
        /// 任职起始时间
        /// </summary>
        public DateTime TakeStartDate { get; set; }
        /// <summary>
        /// 任职结束时间
        /// </summary>
        public DateTime TakEndDate { get; set; }
        /// <summary>
        /// 经营范围
        /// </summary>
        public String Scope { get; set; }
        
        public DateTime CreateTime { get; set; }
    }
}
