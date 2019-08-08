using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFCoreDemo
{
    public class WorkUnit
    {
        public int Id { get; set; }
        /// <summary>
        /// 工作单位名称
        /// </summary>
        public string UnitName { get; set; }
        /// <summary>
        /// 填报人
        /// </summary>
        public string Filler { get; set; }
        /// <summary>
        /// 填报日期
        /// </summary>
        public DateTime FillDate { get; set; }
        /// <summary>
        /// 联系电话
        /// </summary>
        public string Telephone { get; set; }
        /// <summary>
        /// 是否委托
        /// </summary>
        public bool IsEntrust { get; set; }
        /// <summary>
        /// 是否反馈
        /// </summary>
        public bool IsFeedback { get; set; }

        public DateTime CreateTime { get; set; }
    }
}
