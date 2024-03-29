﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using System.IO;
using System.ComponentModel;

namespace EFCoreDemo
{
    public partial class FrmImportProcess : Form
    {

        public FrmImportProcess()
        {
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
        }

        private bool getFeedbackInfo(string fileName)
        {

            XlsToDb xlsToDb = new XlsToDb();
            DataTable dt = xlsToDb.ConvertToDataTable(fileName);
            DataTable dt2 = new DataTable();
            dt2 = dt.Clone();
            List<FeedbackBusiness> feedbackBusinesses = new List<FeedbackBusiness>();
            FeedbackBusiness feedbackBusiness;
            try
            {
                backgroundWorker1.ReportProgress(0, "共检索到" + (dt.Rows.Count-1).ToString() + "家企业信息。\r\n\r\n");
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];
                    feedbackBusiness = new FeedbackBusiness();
                    feedbackBusiness.FullName = dr[0].ToString();
                    feedbackBusiness.CardNo = dr[1].ToString();
                    feedbackBusiness.BusinessName = dr[2].ToString();
                    backgroundWorker1.ReportProgress(0, "正在导入" + feedbackBusiness.BusinessName + "信息。\r\n\r\n");
                    feedbackBusiness.BusinessType = dr[3].ToString();
                    if (!(dr[4] == null || dr[4].ToString().Length == 0))
                    {
                        feedbackBusiness.EstDate = DateTime.Parse(dr[4].ToString());
                    }
                    else
                    {
                        feedbackBusiness.EstDate = DateTime.MinValue;
                    }
                    feedbackBusiness.LegalRep = dr[5].ToString();
                    feedbackBusiness.Residence = dr[6].ToString();
                    feedbackBusiness.State = dr[7].ToString();
                    if (dr[8].ToString().Equals("是"))
                    {
                        feedbackBusiness.IsAbnormal = true;
                    }
                    else
                    {
                        feedbackBusiness.IsAbnormal = false;
                    }
                    if (dr[9].ToString().Equals("是"))
                    {
                        feedbackBusiness.IsOutrage = true;
                    }
                    else
                    {
                        feedbackBusiness.IsOutrage = false;
                    }
                    if (!(dr[10] == null || dr[10].ToString().Length == 0))
                    {
                        feedbackBusiness.RevokeDate = DateTime.Parse(dr[10].ToString());
                    }
                    else
                    {
                        feedbackBusiness.RevokeDate = DateTime.MinValue;
                    }
                    if (!(dr[11] == null || dr[11].ToString().Length == 0))
                    {
                        feedbackBusiness.LogoffDate = DateTime.Parse(dr[11].ToString());
                    }
                    else
                    {
                        feedbackBusiness.LogoffDate = DateTime.MinValue;
                    }
                    feedbackBusiness.CreditCode = dr[12].ToString();
                    if (!(dr[13] == null || dr[13].ToString().Length == 0))
                    {
                        feedbackBusiness.RegisteredCapital = decimal.Parse(dr[13].ToString());
                    }
                    else
                    {
                        feedbackBusiness.RegisteredCapital = 0;
                    }
                    feedbackBusiness.Currency = dr[14].ToString();
                    if (!(dr[15] == null || dr[15].ToString().Length == 0))
                    {
                        feedbackBusiness.Subscribe = decimal.Parse(dr[15].ToString());
                    }
                    else
                    {
                        feedbackBusiness.Subscribe = 0;
                    }
                    if (!(dr[16] == null || dr[16].ToString().Length == 0))
                    {
                        feedbackBusiness.Proportion = decimal.Parse(dr[16].ToString());
                    }
                    else
                    {
                        feedbackBusiness.Proportion = 0;
                    }
                    if (!(dr[17] == null || dr[17].ToString().Length == 0))
                    {
                        feedbackBusiness.PutupDate = DateTime.Parse(dr[17].ToString());
                    }
                    else
                    {
                        feedbackBusiness.PutupDate = DateTime.MinValue;
                    }
                    feedbackBusiness.BusinessPost = dr[18].ToString();
                    if (!(dr[19] == null || dr[19].ToString().Length == 0))
                    {
                        feedbackBusiness.TakeStartDate = DateTime.Parse(dr[19].ToString());
                    }
                    else
                    {
                        feedbackBusiness.TakeStartDate = DateTime.MinValue;
                    }
                    if (!(dr[20] == null || dr[20].ToString().Length == 0))
                    {
                        feedbackBusiness.TakEndDate = DateTime.Parse(dr[20].ToString());
                    }
                    else
                    {
                        feedbackBusiness.TakEndDate = DateTime.MinValue;
                    }
                    feedbackBusiness.Scope = dr[21].ToString();
                    feedbackBusiness.CreateTime = DateTime.Now.Date;
                    feedbackBusinesses.Add(feedbackBusiness);
                    if (i % 20 == 0)
                    {
                        backgroundWorker1.ReportProgress(0, "正在将以上20家企业信息保存至数据库。\r\n\r\n");
                        using (BusinessContext context = new BusinessContext())
                        {
                            //context.Database.EnsureCreated();
                            context.AddRange(feedbackBusinesses);
                            context.SaveChanges();
                        }
                        feedbackBusinesses = new List<FeedbackBusiness>();
                    }
                }
                if (feedbackBusinesses.Count > 0)
                {
                    using (BusinessContext context = new BusinessContext())
                    {
                        //context.Database.EnsureCreated();
                        context.AddRange(feedbackBusinesses);
                        context.SaveChanges();
                    }
                }
                return true;
            }
            catch(Exception ex)
            {
                backgroundWorker1.ReportProgress(0, "出错：" + ex.Message + "\r\n\r\n");
                return false;
            }
                    
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
            int orderId = 1;
            try
            {
                workUnit.UnitName = dt.Rows[1][3].ToString();

                if (workUnit.UnitName.Equals(""))
                {
                    throw new Exception("填报单位不能为空！");
                }
                
                for (int i = 3; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];//9  13

                    if (!(dr[0].ToString().Equals("") && dr[9].ToString().Equals("") && dr[13].ToString().Equals("")))
                    {
                        if (dr[1].ToString().Replace(" ","").Equals("姓名"))
                        {
                            continue;
                        }

                        //如果表格第1列不为空，说明有领导干部信息
                        if (!dr[0].ToString().Equals(""))
                        {
                            if (dr[0].ToString().IndexOf("填报人") > -1)
                            {
                                workUnit.Filler = dr[3].ToString();
                                workUnit.Telephone = dr[15].ToString();
                                workUnit.FillDate = DateTime.Parse(dr[22].ToString());
                                break;
                            }
                            else
                            {
                                if (!dr[1].ToString().Equals(""))
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
                                    leader.FullName = dr[1].ToString().Replace(" ", "").Replace("　", "");
                                    leader.CardNo = dr[2].ToString();
                                    backgroundWorker1.ReportProgress(0, "正在导入领导干部" + leader.FullName + "信息。\r\n\r\n");

                                    leader.Post = dr[3].ToString();
                                    leader.Rank = dr[4].ToString();
                                    if (dr[5].ToString().Equals("存在"))
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
                                    orderId = 1;
                                    familyMember.OrderId = orderId;
                                    orderId++;

                                    familyMember.FullName = leader.FullName;
                                    familyMember.CardNo = leader.CardNo;
                                    familyMember.Relation = "本人";
                                    familyMember.WorkUnit = "本单位";
                                    familyMember.IsValid = true;
                                    if (!IDCardValidation.CheckIDCard(familyMember.CardNo))
                                    {
                                        backgroundWorker1.ReportProgress(0, familyMember.FullName + "身份证号码未通过验证。\r\n\r\n");
                                        familyMember.IsValid = false;
                                    }

                                    if (leader.FamilyMembers == null)
                                    {
                                        leader.FamilyMembers = new List<FamilyMember>();
                                    }
                                    leader.FamilyMembers.Add(familyMember);
                                    familyMember = new FamilyMember();
                                }                                
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

                            
                            familyMember.OrderId = orderId;
                            orderId++;
                            familyMember.Relation = dr[8].ToString();
                            familyMember.FullName = dr[9].ToString();
                            familyMember.CardNo = dr[10].ToString();
                            familyMember.IsValid = true;
                            backgroundWorker1.ReportProgress(0, "正在导入" + familyMember.Relation + familyMember.FullName + "信息。\r\n\r\n");
                            if (!IDCardValidation.CheckIDCard(familyMember.CardNo))
                            {
                                backgroundWorker1.ReportProgress(0, familyMember.FullName + "身份证号码未通过验证。\r\n\r\n");
                                familyMember.IsValid = false;
                            }
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
                            reportBusiness.Remarks = dr[25].ToString();
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
               
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(0, "出错：" + ex.Message+"\r\n\r\n");
                return null;
            }

            return null;
        }

        private void FrmImportProcess_Load(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                // Start the asynchronous operation.
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void BackgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            string targetDir = Directory.GetCurrentDirectory() + "\\导入目录\\待分拣";
            string moveDir = Directory.GetCurrentDirectory() + "\\导入目录\\已分拣\\";
            List<FileInfo> fileInformations = DirectoryAllFiles.GetAllFiles(targetDir);
            worker.ReportProgress(0, "共找到" + fileInformations.Count.ToString() + "个文件\r\n\r\n");
            foreach (FileInfo fi in fileInformations)
            {
                worker.ReportProgress(0, "正在导入" + fi.Name + "文件。\r\n\r\n");
                if (fi.Name.IndexOf("工商反馈") > -1)
                {
                    if (getFeedbackInfo(fi.FullName))
                    {
                        worker.ReportProgress(0, fi.Name + "导入完成；\r\n\r\n");
                        fi.MoveTo(moveDir + fi.Name);
                        worker.ReportProgress(0, "移动" + fi.Name + "文件至已分拣目录\r\n\r\n");
                    }
                    //List<FeedbackBusiness> feedbackBusinesses = getFeedbackInfO(fi.FullName);
                    //if (feedbackBusinesses.Count>0)
                    //{
                    //    using (BusinessContext context = new BusinessContext())
                    //    {                            
                    //        //context.Database.EnsureCreated();
                    //        context.AddRange(feedbackBusinesses);
                    //        context.SaveChanges();
                    //    }
                    //    if (File.Exists(moveDir + fi.Name))
                    //    {
                    //        File.Delete(moveDir + fi.Name);
                    //    }

                    //    worker.ReportProgress(0,fi.Name+ "导入完成；\r\n\r\n");
                    //    fi.MoveTo(moveDir + fi.Name);
                    //    worker.ReportProgress(0, "移动" + fi.Name + "文件至已分拣目录\r\n\r\n");
                    //}
                    else
                    {
                        worker.ReportProgress(0, "导入" + fi.Name + "失败，请检查文件内容。\r\n\r\n");
                    }
                }
                else
                {
                    WorkUnit workUnit = getCollectInfo(fi.FullName);
                    if (workUnit != null)
                    {
                        using (BusinessContext context = new BusinessContext())
                        {
                            var dworkUnit = context.WorkUnits.Include(b => b.Leaders).Where(c => c.UnitName == workUnit.UnitName);
                            if (dworkUnit.Count() > 0)
                            {
                                WorkUnit work = dworkUnit.First();

                                //List<Leader> leaders = context.Leaders.Where(b => b.WorkUnit == dworkUnit).ToList();
                                //context.Leaders.RemoveRange(leaders);
                                context.WorkUnits.Remove(work);
                                context.SaveChanges();
                            }
                            //context.Database.EnsureCreated();
                            context.Add(workUnit);
                            context.SaveChanges();
                        }
                        if (File.Exists(moveDir + fi.Name))
                        {
                            File.Delete(moveDir + fi.Name);
                        }

                        int familyCount = 0;
                        foreach (Leader leader in workUnit.Leaders)
                        {
                            familyCount += leader.FamilyMembers.Count;
                        }

                        worker.ReportProgress(0, "导入完成，" + workUnit.UnitName + "共导入" + workUnit.Leaders.Count + "位领导干部信息、" +
                            familyCount + "位家庭成员信息（含领导本人）；\r\n\r\n");
                        fi.MoveTo(moveDir + fi.Name);
                        worker.ReportProgress(0, "移动" + fi.Name + "文件至已分拣目录\r\n\r\n");
                    }
                    else
                    {
                        worker.ReportProgress(0, "导入" + fi.Name + "失败，请检查文件内容。\r\n\r\n");
                    }
                }                
                
            }
           worker.ReportProgress(0, "导入完成，请关闭窗口\r\n");
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            textBox1.Text += e.UserState.ToString();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Select(textBox1.Text.Length, 0);
            textBox1.ScrollToCaret();
        }
    }
    
}
