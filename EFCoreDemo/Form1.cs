using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            Leader leader;
            FamilyMember familyMember;
            List<Leader> leaders = new List<Leader>();
            for (int i = 4; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];//9  13                

                if (!(dr[0].ToString().Equals("") && dr[9].ToString().Equals("") && dr[13].ToString().Equals("")))
                {
                    if (!dr[0].ToString().Equals(""))
                    {
                        leader = new Leader();
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

                        if (!dr[9].ToString().Equals(""))
                        {
                            familyMember = new FamilyMember();
                            familyMember.Relation =
                        }
                    }
                    else
                    {

                    }


                    dt2.ImportRow(dr);
                    leaders.Add(leader);
                }
            }
            dataGridView1.DataSource = leaders;


        }

    }
}
