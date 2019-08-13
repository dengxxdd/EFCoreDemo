﻿using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.EntityFrameworkCore;
using System.IO;
using System.Threading;

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
            //显示进度窗体
            FrmImportProcess frm = new FrmImportProcess();
            frm.StartPosition = FormStartPosition.CenterScreen;
            frm.ShowDialog(this);
        }
 

        private void Button2_Click(object sender, EventArgs e)
        {
            
        }

        private WorkUnit getCollectInfo(string fileName)
        {
            XlsToDb xlsToDb = new XlsToDb();
            DataTable dt = xlsToDb.ConvertToDataTable(fileName);
            DataTable dt2 = new DataTable();
            dt2 = dt.Clone();
            Leader leader = new Leader();
            List<Leader> leaders = new List<Leader>();
            FamilyMember familyMember = new FamilyMember();
            ReportBusiness reportBusiness = new ReportBusiness();
            WorkUnit workUnit = new WorkUnit();
            try
            {
                workUnit.UnitName = dt.Rows[1][2].ToString();

                if (workUnit.UnitName.Equals(""))
                {
                    throw new Exception("填报单位不能为空！");
                }

                workUnit.FillDate = DateTime.Parse(dt.Rows[1][21].ToString());
                for (int i = 3; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];//9  13

                    if (!(dr[0].ToString().Equals("") && dr[9].ToString().Equals("") && dr[13].ToString().Equals("")))
                    {

                        //如果表格第1列不为空，说明有领导干部信息
                        if (!dr[0].ToString().Equals(""))
                        {
                            if (dr[0].ToString().IndexOf("填报人") > -1)
                            {
                                workUnit.Filler = dr[2].ToString();
                                workUnit.Telephone = dr[18].ToString();
                            }
                            else
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

                workUnit.Leaders = leaders;

                return workUnit;

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

            List<FileInfo> fileInformations = DirectoryAllFiles.GetAllFiles(targetDir);
            foreach (FileInfo fi in fileInformations)
            {
                WorkUnit workUnit = getCollectInfo(fi.FullName);
                if (workUnit != null)
                {
                    using (BusinessContext context = new BusinessContext())
                    {
                        //var dworkUnit = context.WorkUnits.Include(b => b.Leaders).Single(b => b.UnitName == workUnit.UnitName);
                        //if (dworkUnit != null)
                        //{
                        //    List<Leader> leaders = context.Leaders.Where(b => b.WorkUnit == dworkUnit).ToList();
                        //    context.Leaders.RemoveRange(leaders);
                        //    context.WorkUnits.Remove(dworkUnit);
                        //    context.SaveChanges();
                        //}
                        context.Database.EnsureCreated();
                        context.AddRange(workUnit);
                        context.SaveChanges();
                    }
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
                               BusinessId = eci.Id.ToString(),
                               eci.BusinessName,
                               eci.BusinessType
                           };
                


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

                var data3 = from a in context.WorkUnits
                            join b in context.Leaders
                            on a.Id equals b.WorkUnitId
                            join c in context.FamilyMembers
                            on b.Id equals c.LeaderId
                            where (a.IsEntrust == false)
                            select new
                            {
                                LeaderId = c.LeaderId,
                                FullName = c.FullName,
                                CardNo = c.CardNo,
                                ReportDate = DateTime.Now
                           };

                #endregion

                var workUnits = context.WorkUnits.Include(a=>a.Leaders).ThenInclude(b=>b.FamilyMembers).Where(a => a.IsEntrust == false).ToList();
                List<FamilyMember> familyMembers = new List<FamilyMember>();
                foreach(WorkUnit workUnit in workUnits)
                {
                    foreach(Leader leader in workUnit.Leaders)
                    {
                        familyMembers.AddRange(leader.FamilyMembers);
                    }
                    workUnit.IsEntrust = true;
                }
                CollectToFile( familyMembers);
                context.UpdateRange(workUnits);
                context.SaveChanges();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            decimal a = Convert.ToDecimal("10");
            int b = 0;
            //    var q = from c in categories
            //join p in products on c equals p.Category into ps
            //from p in ps.DefaultIfEmpty()
            //select new { Category = c, ProductName = p == null ? "(No products)" : p.ProductName };
        }

        public int[] NPOI_CellChange(string A1)
        {
            int[] nCell = new int[2];
            string strSplit1, strSplit2;
            //string A2 = A1.Substring(7, 3);
            strSplit1 = Regex.Replace(A1, "[0-9]", "", RegexOptions.IgnoreCase);
            strSplit2 = Regex.Replace(A1, "[a-z]", "", RegexOptions.IgnoreCase);
            char[] A3 = strSplit1.ToCharArray();
            nCell[0] = int.Parse(strSplit2) - 1;

            int n = 0;
            for (int i = 0; i < A3.Length; i++)
            {
                n += (A3[i] - 'A' + 1) * Convert.ToInt32((System.Math.Pow(26, Convert.ToDouble(A3.Length - i - 1))));
            }
            nCell[1] = n - 1;
            return nCell;
        }        

        //将比对结果添加到汇总表         
        public string CollectToFile( List<FamilyMember> familyMembers)
        {
            int rowIndex = 0;   //excel新建行的行号
            IWorkbook workbook;
            IRow row;
            ICell cell;
            ICellStyle style, style2;
            //HSSFCellStyle celStyle = getCellStyle();

            try
            {
                using (FileStream file = new FileStream(System.Environment.CurrentDirectory + "\\Templates\\工商查询模板.xlsx", FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(file);
                }

                if (workbook is HSSFWorkbook)
                {
                    style = ((HSSFWorkbook)workbook).CreateCellStyle();
                    style2 = ((HSSFWorkbook)workbook).CreateCellStyle();
                }
                else
                {
                    style = ((XSSFWorkbook)workbook).CreateCellStyle();
                    style2 = ((XSSFWorkbook)workbook).CreateCellStyle();
                }

                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BottomBorderColor = 128;
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                style.VerticalAlignment = VerticalAlignment.Center;
                style.WrapText = true;
                style2.CloneStyleFrom(style);
                style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                ISheet sheet = workbook.GetSheetAt(0); 

                for (int i = 0; i < familyMembers.Count; i++)
                {
                    rowIndex = 3 + i;
                    row = sheet.CreateRow(rowIndex);
                    for (int j = 0; j < 4; j++)
                    {
                        cell = row.CreateCell(j);
                        cell.CellStyle = style;
                    }

                    cell = sheet.GetRow(rowIndex).GetCell(0);
                    cell.SetCellValue(i + 1);

                    cell = sheet.GetRow(rowIndex).GetCell(1);
                    cell.SetCellValue(familyMembers[i].FullName);
                    cell = sheet.GetRow(rowIndex).GetCell(2);
                    cell.SetCellValue(familyMembers[i].CardNo);
                    cell = sheet.GetRow(rowIndex).GetCell(3);
                    cell.SetCellValue(DateTime.Now.ToString("yyyyMMdd"));
                }

                string targetDir = Directory.GetCurrentDirectory() + "\\委托查询\\" + DateTime.Now.ToString("yyyyMMdd") + "工商查询.xlsx";
                if (File.Exists(targetDir))
                {
                    File.Delete(targetDir);
                }
                using (FileStream fileStream = File.Open(targetDir, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileStream);
                    fileStream.Close();
                }
                workbook.Close();
                return "OK";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

    }
}
