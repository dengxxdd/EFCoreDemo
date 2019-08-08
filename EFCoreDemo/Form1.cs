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
using System.IO;

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

        private List<Leader> getCollectInfo(string fileName)
        {
            XlsToDb xlsToDb = new XlsToDb();
            DataTable dt = xlsToDb.ConvertToDataTable(fileName);
            DataTable dt2 = new DataTable();
            dt2 = dt.Clone();
            Leader leader = new Leader();
            List<Leader> leaders = new List<Leader>();
            FamilyMember familyMember = new FamilyMember();
            ReportBusiness reportBusiness = new ReportBusiness();

            try
            {
                for (int i = 3; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];//9  13

                    if (!(dr[0].ToString().Equals("") && dr[9].ToString().Equals("") && dr[13].ToString().Equals("")))
                    {

                        //如果表格第1列不为空，说明有领导干部信息
                        if (!dr[0].ToString().Equals(""))
                        {
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
                                leader = new Leader();
                            }
                            leader.OrderId = int.Parse(dr[0].ToString());
                            leader.FullName = dr[1].ToString();
                            leader.CardNo = dr[2].ToString();

                            //if (!IDCardValidation.CheckIDCard(leader.CardNo))
                            //{
                            //    throw new Exception("导入'经商办企业情况汇总表.xlsx'时出错：" + leader.FullName + "身份证校验未出错");
                            //}

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

                            familyMember = new FamilyMember();
                            familyMember.FullName = leader.FullName;
                            familyMember.CardNo = leader.CardNo;
                            familyMember.Relation = "本人";
                            familyMember.WorkUnit = "本单位";

                            if (leader.FamilyMembers == null)
                            {
                                leader.FamilyMembers = new List<FamilyMember>();
                            }
                            leader.FamilyMembers.Add(familyMember);
                            familyMember = new FamilyMember();
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

                                if (leader.FamilyMembers == null)
                                {
                                    leader.FamilyMembers = new List<FamilyMember>();
                                }
                                leader.FamilyMembers.Add(familyMember);
                                familyMember = new FamilyMember();
                            }



                            familyMember.Relation = dr[8].ToString();
                            familyMember.FullName = dr[9].ToString();
                            familyMember.CardNo = dr[10].ToString();
                            //if (!IDCardValidation.CheckIDCard(familyMember.CardNo))
                            //{
                            //    throw new Exception("导入'经商办企业情况汇总表.xlsx'时出错："+ familyMember.FullName + "身份证校验未出错");
                            //}
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
                                reportBusiness.Subscribe = decimal.Parse(dr[17].ToString());
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
                                reportBusiness.EstDate = DateTime.Parse(dr[20].ToString());
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

                return leaders;

                //using (BusinessContext context = new BusinessContext())
                //{
                //    context.Database.EnsureCreated();
                //    context.AddRange(leaders);
                //    context.SaveChanges();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

            return null;
        }

        private void Button3_Click(object sender, EventArgs e)
        {                  
            string targetDir = Directory.GetCurrentDirectory() + "\\导入目录\\待分拣";
            
            List<FileInfo> fileInformations= DirectoryAllFiles.GetAllFiles(targetDir);
            foreach(FileInfo fi in fileInformations)
            {
                List<Leader> leaders = getCollectInfo(fi.FullName);
                using (BusinessContext context = new BusinessContext())
                {
                    context.Database.EnsureCreated();
                    context.AddRange(leaders);
                    context.SaveChanges();
                }
            }            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (BusinessContext context = new BusinessContext())
            {

                #region 多级联合查询:所有经商办企业联合查询
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
                #endregion


                var data2 = from a in context.Leaders
                           join b in context.FamilyMembers
                           on a.Id equals b.LeaderId into dc
                           from dci in dc.DefaultIfEmpty()
                           select new
                           {
                               LeaderId = a.Id,
                               LeaderFullName = a.FullName,
                               LeaderCardNo = a.CardNo,
                               FamilyId = dci.Id.ToString(),
                               dci.Relation,
                               FamilyFullName = dci.FullName,
                               FamilyCardNo = dci.CardNo                               
                           };

                dataGridView1.DataSource = data2.ToList();
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
