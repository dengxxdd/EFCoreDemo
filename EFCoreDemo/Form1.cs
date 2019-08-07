using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.EntityFrameworkCore;

namespace EFCoreDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            using(BloggingContext context=new BloggingContext())
            {
                //context.Database.EnsureCreated();
                List<Blog> blogs = new List<Blog>();
                Blog current;
                for(int i = 1; i <= 10; i++)
                {
                    current = context.Blogs.Single(b => b.BlogId == i);
                    current.Url = "texts" + i.ToString();
                    blogs.Add(current);
                }
                context.UpdateRange(blogs);
                context.SaveChanges();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            using (BloggingContext context = new BloggingContext())
            {
                var blogs = context.Blogs.Where(b=>new int?[] { 1, 2,5,8 }.Contains(b.BlogId)).ToList();
                dataGridView1.DataSource = blogs;
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            XlsToDb xlsToDb = new XlsToDb();
            DataTable dt = xlsToDb.ConvertToDataTable("经商办企业情况汇总表.xlsx");
            DataTable dt2 = new DataTable();
            dt2 = dt.Clone();
            Leader leader = new Leader();            
            List<Leader> leaders = new List<Leader>();
            FamilyMember familyMember = new FamilyMember();
            ReportBusiness reportBusiness = new ReportBusiness();

            #region 注释掉的代码
            //leader.OrderId = 1;
            //leader.FullName = "张三";
            //leader.CardNo = "431111197602040000";
            //leader.Post = "书记";
            //leader.Rank = "正处";
            //leader.IsExist = true;
            //leader.IsRegulated = false;
            //leader.FamilyMembers = new List<FamilyMember>();


            //familyMember.ReportBusinesses = new List<ReportBusiness>();
            //familyMember.Relation = "配偶";
            //familyMember.FullName = "李四";
            //familyMember.CardNo = "431111197803050021";
            //familyMember.WorkUnit = "教育"; 
            //leader.FamilyMembers.Add(familyMember);

            //reportBusiness = new ReportBusiness();
            //reportBusiness.BusinessType = "A1";
            //reportBusiness.BusinessName = "有限公司";
            //reportBusiness.BusinessPost = "执行董事";
            //reportBusiness.RegPlace = "长沙";
            //reportBusiness.BusinessPlace = "长沙";
            //reportBusiness.Subscribe = 200;
            //reportBusiness.PaidinCapital = 20;
            //reportBusiness.Proportion = 50;
            //reportBusiness.EstDate = DateTime.Parse("2017-08-02");
            //reportBusiness.Nearly3Years1 = "百货、餐饮";
            //reportBusiness.Nearly3Years2 = "XX公司";
            //reportBusiness.Nearly3Years3 = "";
            //reportBusiness.IsRelevant = false;
            //familyMember.ReportBusinesses.Add(reportBusiness);

            //reportBusiness = new ReportBusiness();
            //reportBusiness.BusinessType = "A2";
            //reportBusiness.BusinessName = "股份公司";
            //reportBusiness.BusinessPost = "股东";
            //reportBusiness.RegPlace = "广州";
            //reportBusiness.BusinessPlace = "广州";
            //reportBusiness.Subscribe = 100;
            //reportBusiness.PaidinCapital = 10;
            //reportBusiness.Proportion = 30;
            //reportBusiness.EstDate = DateTime.Parse("2013-02-02");
            //reportBusiness.Nearly3Years1 = "住宿、娱乐";
            //reportBusiness.Nearly3Years2 = "YY公司";
            //reportBusiness.Nearly3Years3 = "";
            //reportBusiness.IsRelevant = false;
            //familyMember.ReportBusinesses.Add(reportBusiness);


            //familyMember = new FamilyMember();
            //familyMember.Relation = "儿子";
            //familyMember.FullName = "张四";
            //familyMember.CardNo = "431111200201030019";
            //familyMember.WorkUnit = "城管";
            //leader.FamilyMembers.Add(familyMember);
            //familyMember = new FamilyMember();
            //familyMember.Relation = "儿媳";
            //familyMember.FullName = "李五";
            //familyMember.CardNo = "431111200508090020";
            //familyMember.WorkUnit = "文化";




            //leader.FamilyMembers.Add(familyMember);

            //leaders.Add(leader);

            //using (BusinessContext context = new BusinessContext())
            //{
            //    context.Database.EnsureCreated();
            //    //context.Database.EnsureCreated();
            //    context.AddRange(leader);
            //    context.SaveChanges();
            //}

            //dataGridView1.DataSource = leaders;
            #endregion

            for (int i = 3; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];//9  13

                if (!(dr[0].ToString().Equals("") && dr[9].ToString().Equals("") && dr[13].ToString().Equals("")))
                {

                    //如果表格第1列不为空，说明有领导干部信息
                    if (!dr[0].ToString().Equals(""))
                    {
                        if(!(leader.CardNo == null || leader.CardNo.Length == 0))
                        {
                            if (!(reportBusiness.BusinessName == null || reportBusiness.BusinessName.Length == 0))
                            {
                                if (familyMember.ReportBusinesses == null)
                                {
                                    familyMember.ReportBusinesses = new List<ReportBusiness>();
                                }
                                familyMember.ReportBusinesses.Add(reportBusiness);
                                reportBusiness = new ReportBusiness();
                            }

                            if (!(familyMember.CardNo == null || familyMember.CardNo.Length == 0))
                            {
                                if (leader.FamilyMembers == null)
                                {
                                    leader.FamilyMembers = new List<FamilyMember>();
                                }
                                leader.FamilyMembers.Add(familyMember);
                                familyMember = new FamilyMember();
                            }                            

                            leaders.Add(leader);
                            leader = new Leader();
                        }
                        leader.OrderId = int.Parse(dr[0].ToString());
                        leader.FullName = dr[1].ToString();
                        leader.CardNo = dr[2].ToString();
                        leader.Post = dr[3].ToString();
                        leader.Rank = dr[4].ToString();
                        if (dr[5].ToString().Equals("是"))
                        {
                            leader.IsExist = true;
                        }
                        else
                        {
                            leader.IsExist = false;
                        }
                        if (dr[6].ToString().Equals("是"))
                        {
                            leader.IsRegulated = true;
                        }
                        else
                        {
                            leader.IsRegulated = false;
                        }

                        if (!dr[7].ToString().Equals(""))
                        {
                            leader.ExitModel = dr[7].ToString();
                        }
                    }

                    if (!dr[9].ToString().Equals(""))
                    {
                        if (!(familyMember.CardNo == null || familyMember.CardNo.Length == 0))
                        {
                            if (!(reportBusiness.BusinessName == null || reportBusiness.BusinessName.Length == 0))
                            {
                                if (familyMember.ReportBusinesses == null)
                                {
                                    familyMember.ReportBusinesses = new List<ReportBusiness>();
                                }
                                familyMember.ReportBusinesses.Add(reportBusiness);
                                reportBusiness = new ReportBusiness();
                            }

                            if (leader.FamilyMembers == null) {
                                leader.FamilyMembers = new List<FamilyMember>();
                            }
                            leader.FamilyMembers.Add(familyMember);
                            familyMember = new FamilyMember();
                        }


                        
                        familyMember.Relation = dr[8].ToString();
                        familyMember.FullName = dr[9].ToString();
                        familyMember.CardNo = dr[10].ToString();
                        familyMember.WorkUnit = dr[11].ToString();
                    }

                    if (!dr[13].ToString().Equals(""))
                    {
                        if (!(reportBusiness.BusinessName == null || reportBusiness.BusinessName.Length == 0))
                        {
                            if (familyMember.ReportBusinesses == null)
                            {
                                familyMember.ReportBusinesses = new List<ReportBusiness>();
                            }
                            familyMember.ReportBusinesses.Add(reportBusiness);
                            reportBusiness = new ReportBusiness();
                        }

                        
                        reportBusiness.BusinessType = dr[12].ToString();
                        reportBusiness.BusinessName = dr[13].ToString();
                        reportBusiness.RegPlace = dr[14].ToString();
                        reportBusiness.BusinessPlace = dr[15].ToString();
                        reportBusiness.BusinessPost = dr[16].ToString();

                        if (dr[17].ToString().Equals(""))
                        {
                            reportBusiness.Subscribe = 0;
                        }
                        else
                        {
                            reportBusiness.Subscribe=decimal.Parse(dr[17].ToString());
                        }
                        if (dr[18].ToString().Equals(""))
                        {
                            reportBusiness.PaidinCapital = 0;
                        }
                        else
                        {
                            reportBusiness.PaidinCapital = decimal.Parse(dr[18].ToString());
                        }
                        if (dr[19].ToString().Equals(""))
                        {
                            reportBusiness.Proportion = 0;
                        }
                        else
                        {
                            reportBusiness.Proportion = decimal.Parse(dr[19].ToString());
                        }

                        if (dr[20].ToString().Equals(""))
                        {
                            reportBusiness.EstDate = DateTime.MinValue;
                        }
                        else
                        {
                            reportBusiness.EstDate =DateTime.Parse(dr[20].ToString());
                        }
                        reportBusiness.Nearly3Years1 = dr[21].ToString();
                        reportBusiness.Nearly3Years2 = dr[22].ToString();
                        reportBusiness.Nearly3Years3 = dr[23].ToString();
                        if (dr[24].ToString().Equals("是"))
                        {
                            reportBusiness.IsRelevant = true;
                        }
                        else
                        {
                            reportBusiness.IsRelevant = false;
                        }                        
                    }
                }

            }

            if (!(leader.CardNo == null || leader.CardNo.Length == 0))
            {
                if (!(reportBusiness.BusinessName == null || reportBusiness.BusinessName.Length == 0))
                {
                    if (familyMember.ReportBusinesses == null)
                    {
                        familyMember.ReportBusinesses = new List<ReportBusiness>();
                    }
                    familyMember.ReportBusinesses.Add(reportBusiness);
                    reportBusiness = new ReportBusiness();
                }

                if (!(familyMember.CardNo == null || familyMember.CardNo.Length == 0))
                {
                    if (leader.FamilyMembers == null)
                    {
                        leader.FamilyMembers = new List<FamilyMember>();
                    }
                    leader.FamilyMembers.Add(familyMember);
                    familyMember = new FamilyMember();
                }
                leaders.Add(leader);                
            }

            dataGridView1.DataSource = leaders.ToList();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (BusinessContext context = new BusinessContext())
            {
                
                //多级联合查询
                var data = from a in context.Leaders
                           join b in context.FamilyMembers
                           on a.Id equals b.LeaderId into dc
                           from dci in dc.DefaultIfEmpty()
                           join c in context.ReportBusinesses
                           on dci.Id equals c.FamilyMemberId into ec
                           from eci in ec.DefaultIfEmpty()
                           select new
                           {
                               LeaderId = a.Id,
                               LeaderFullName = a.FullName,
                               LeaderCardNo = a.CardNo,
                               FamilyId = dci.Id.ToString(),
                               dci.Relation,
                               FamilyFullName = dci.FullName,
                               FamilyCardNo = dci.CardNo,
                               BusinessId=eci.Id.ToString(),
                               eci.BusinessName,
                               eci.BusinessType
                           };
                dataGridView1.DataSource = data.ToList();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            decimal a=Convert.ToDecimal("10");
            int b = 0;
        //    var q = from c in categories
        //join p in products on c equals p.Category into ps
        //from p in ps.DefaultIfEmpty()
        //select new { Category = c, ProductName = p == null ? "(No products)" : p.ProductName };
        }
    }
}
