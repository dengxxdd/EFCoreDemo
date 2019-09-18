using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
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
            
        }
 

        private void Button2_Click(object sender, EventArgs e)
        {
            
        }

        
        private void Button3_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
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

        public string CompareToFile(Leader leader)
        {
            Dictionary<string, string> templateStr = new Dictionary<string, string>();
            templateStr["reportInvest"] = "{Relation+Name}投资{BusinessName},认缴出资{Subscribe}万元，出资比例{Proportion}%。";
            templateStr["feedbackInvest"] = "{Relation+Name}投资{BusinessName},注册资本{RegisteredCapital}万元,认缴出资{Subscribe}万元，出资比例{Proportion}%。";
            templateStr["Post"] = "{Relation+Name}在{BusinessName}担任{Post}职务";


            int rowIndex = 0;   //excel新建行的行号
            IWorkbook workbook;
            IRow row;
            ICell cell;
            ICellStyle style, style2;
            //HSSFCellStyle celStyle = getCellStyle();

            try
            {
                using (FileStream file = new FileStream(System.Environment.CurrentDirectory + "\\Templates\\比对模板.xlsx", FileMode.Open, FileAccess.Read))
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

                int[] pos = NPOI_CellChange("C3");
                cell = sheet.GetRow(pos[0]).GetCell(pos[1]);
                cell.SetCellValue(leader.FullName);
                pos = NPOI_CellChange("F3");
                cell = sheet.GetRow(pos[0]).GetCell(pos[1]);
                cell.SetCellValue(leader.WorkUnit.UnitName);
                pos = NPOI_CellChange("C4");
                cell = sheet.GetRow(pos[0]).GetCell(pos[1]);
                cell.SetCellValue(leader.Post);
                pos = NPOI_CellChange("F4");
                cell = sheet.GetRow(pos[0]).GetCell(pos[1]);
                cell.SetCellValue(leader.Rank);
                pos = NPOI_CellChange("C5");
                cell = sheet.GetRow(pos[0]).GetCell(pos[1]);
                if (leader.IsExist)
                {
                    cell.SetCellValue("存在");
                }
                else
                {
                    cell.SetCellValue("不存在");
                }

                pos = NPOI_CellChange("E5");
                cell = sheet.GetRow(pos[0]).GetCell(pos[1]);
                if (leader.IsRegulated)
                {
                    cell.SetCellValue("是");
                }
                else
                {
                    cell.SetCellValue("否");
                }
                pos = NPOI_CellChange("G5");
                cell = sheet.GetRow(pos[0]).GetCell(pos[1]);
                if (!(leader.ExitModel == null || leader.ExitModel.Length == 0))
                {
                    cell.SetCellValue(leader.ExitModel);
                }

                //MyInsertRow(sheet, 6, leader.FamilyMembers.Count, row);

                Dictionary<string, List<ReportBusiness>> dicReportBusiness = new Dictionary<string, List<ReportBusiness>>();
                Dictionary<string, List<FeedbackBusiness>> dicFeedbackBusiness = new Dictionary<string, List<FeedbackBusiness>>();

                //reportCount：个人报告的企业总数，feedbackCount：反馈的企业总数，
                //addCount：需增加的表格行数，如果个人报告和反馈企业数都为0，行数为1，否则，行数为个人报告和反馈企业总数值大的那个数
                int reportCount = 0, feedbackCount = 0,addCount=0;

                for (int i = 0; i < leader.FamilyMembers.Count; i++)
                {
                    FamilyMember familyMember = leader.FamilyMembers[i];
                    if (familyMember.Relation != "本人")
                    {
                        rowIndex = 6 + i-1;
                        //row = sheet.CreateRow(rowIndex);
                        sheet.ShiftRows(
                            rowIndex,                            //--开始行
                            sheet.LastRowNum,                      //--结束行
                                                                   //leader.FamilyMembers.Count,                             //--移动大小(行数)--往下移动
                            1,
                            true,                                  //是否复制行高
                            false//,                               //是否重置行高
                        );
                        row = sheet.CreateRow(rowIndex);

                        for (int j = 0; j < 7; j++)
                        {
                            cell = row.CreateCell(j);
                            if (j == 1)     //关系栏单元格居中，其他的靠左
                            {
                                cell.CellStyle = style;
                            }
                            else
                            {
                                cell.CellStyle = style2;
                            }                            
                        }
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 3));
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 6));
                        row.Height = 36 * 20;
                        cell = row.GetCell(1);
                        cell.SetCellValue(familyMember.Relation);
                        cell = row.GetCell(2);
                        cell.SetCellValue(familyMember.FullName);
                        cell = row.GetCell(4);
                        cell.SetCellValue(familyMember.WorkUnit);
                        if (familyMember.ReportBusinesses.Count > 0)
                        {
                            dicReportBusiness[familyMember.Relation + familyMember.FullName] = familyMember.ReportBusinesses;
                            reportCount += familyMember.ReportBusinesses.Count;     //统计个人填报总共有多少家企业，在后面的个人填报情况中判断要增加几行
                        }                        
                        using (BusinessContext context = new BusinessContext())
                        {
                            var feedbacks = context.FeedbackBusinesses.Where(a=>a.CardNo== familyMember.CardNo).ToList();
                            if (feedbacks.Count > 0)
                            {
                                dicFeedbackBusiness[familyMember.Relation+familyMember.FullName] = feedbacks;
                                feedbackCount += feedbacks.Count;       //统计信息反馈总共有多少家企业，在后面的个人填报情况中判断要增加几行
                            }                            
                        }

                    }                    
                }

                sheet.AddMergedRegion(new CellRangeAddress(rowIndex - leader.FamilyMembers.Count+1, rowIndex, 0, 0));
                if (reportCount < feedbackCount)
                {
                    addCount = feedbackCount;
                }
                else
                {
                    addCount = reportCount;
                }
                if (addCount == 0)
                {
                    addCount = 1;
                }
                rowIndex+=2;
                int rowIndex2 = rowIndex;//记录当前行号

                sheet.ShiftRows(
                            rowIndex,                            //--开始行
                            sheet.LastRowNum,                      //--结束行
                                                                   //leader.FamilyMembers.Count,                             //--移动大小(行数)--往下移动
                            addCount,
                            true,                                  //是否复制行高
                            false//,                               //是否重置行高
                        );

                for(int i = 0; i < addCount; i++)
                {
                    row = sheet.CreateRow(rowIndex);
                    for (int j = 0; j < 7; j++)
                    {
                        cell = row.CreateCell(j);
                        cell.CellStyle = style2;
                    }
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 6));
                    row.Height = 50 * 20;
                    rowIndex++;
                }

                sheet.AddMergedRegion(new CellRangeAddress(rowIndex - addCount-1, rowIndex-1, 0, 0));

                rowIndex = rowIndex2;
                reportCount = 1;
                feedbackCount = 1;
                foreach (KeyValuePair<string, List<ReportBusiness>> reports in dicReportBusiness)
                {
                    foreach (ReportBusiness report in reports.Value)
                    {
                        row = sheet.GetRow(rowIndex);                        
                        cell = row.GetCell(1);

                        if (report.Subscribe != 0)
                        {
                            //templateStr["reportInvest"] = "{Relation+Name}投资{BusinessName},认缴出资{Subscribe}万元，出资比例{Proportion}%。";
                            cell.SetCellValue(reportCount.ToString()+"."+templateStr["reportInvest"].Replace("{Relation+Name}", reports.Key).Replace("{BusinessName}", report.BusinessName).Replace("{Subscribe}", report.Subscribe.ToString())
                            .Replace("{Proportion}", (report.Proportion * 100).ToString()));
                        }
                        else
                        {
                            //templateStr["Post"] = "{Relation+Name}在{BusinessName}担任{Post}职务";
                            cell.SetCellValue(reportCount.ToString() + "." + templateStr["Post"].Replace("{Relation+Name}", reports.Key).Replace("{BusinessName}", report.BusinessName).Replace("{Post}", report.BusinessPost));
                        }                        
                        rowIndex++;
                        reportCount++;
                    }
                }

                rowIndex = rowIndex2;
                foreach (KeyValuePair<string, List<FeedbackBusiness>> feedbacks in dicFeedbackBusiness)
                {
                    foreach (FeedbackBusiness feedback in feedbacks.Value)
                    {
                        row = sheet.GetRow(rowIndex);
                        cell = row.GetCell(4);
                        if (feedback.Subscribe != 0)
                        {
                            //templateStr["feedbackInvest"] = "{Relation+Name}投资{BusinessName},注册资本{RegisteredCapital}万元,认缴出资{Subscribe}万元，出资比例{Proportion}%。";
                            cell.SetCellValue(feedbackCount.ToString() + "." + templateStr["feedbackInvest"].Replace("{Relation+Name}", feedbacks.Key).Replace("{BusinessName}", feedback.BusinessName)
                                .Replace("{RegisteredCapital}",feedback.RegisteredCapital.ToString()).Replace("{ Subscribe}", feedback.Subscribe.ToString())
                            .Replace("{Proportion}", (feedback.Proportion * 100).ToString()));
                        }
                        else
                        {
                            //templateStr["Post"] = "{Relation+Name}在{BusinessName}担任{Post}职务";
                            cell.SetCellValue(feedbackCount.ToString() + "." + templateStr["Post"].Replace("{Relation+Name}", feedbacks.Key).Replace("{BusinessName}", feedback.BusinessName).Replace("{Post}", feedback.BusinessPost));
                        }
                        rowIndex++;
                        feedbackCount++;
                    }
                }

                string targetDir = Directory.GetCurrentDirectory() + "\\比对表\\" + leader.OrderId.ToString() +". "+ leader.FullName+".xlsx";
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
                MessageBox.Show(e.Message);
                return e.Message;
            }
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
                cell = sheet.GetRow(0).GetCell(0);
                cell.SetCellValue(DateTime.Now.ToString("yyyyMMdd") + "委托查询（工商）");
                cell = sheet.GetRow(1).GetCell(3);
                cell.SetCellValue(DateTime.Now.ToString("yyyyMMdd"));
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
                MessageBox.Show(e.Message);
                return e.Message;
            }
        }

        private void ToolStripButton1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("43010319841016152X" + IDCardValidation.CheckIDCard("43010319841016152X").ToString());
            //显示进度窗体
            FrmImportProcess frm = new FrmImportProcess();
            frm.StartPosition = FormStartPosition.CenterScreen;
            frm.ShowDialog(this);
        }

        private void ToolStripButton2_Click(object sender, EventArgs e)
        {
            using (BusinessContext context = new BusinessContext())
            {

                #region 多级联合查询:所有经商办企业联合查询
                //var data = from a in context.Leaders
                //           join b in context.FamilyMembers
                //           on a.Id equals b.LeaderId into dc
                //           from dci in dc.DefaultIfEmpty()
                //           join c in context.ReportBusinesses
                //           on dci.Id equals c.FamilyMemberId into ec
                //           from eci in ec.DefaultIfEmpty()
                //           select new
                //           {
                //               LeaderId = a.Id,
                //               LeaderFullName = a.FullName,
                //               LeaderCardNo = a.CardNo,
                //               FamilyId = dci.Id.ToString(),
                //               dci.Relation,
                //               FamilyFullName = dci.FullName,
                //               FamilyCardNo = dci.CardNo,
                //               BusinessId = eci.Id.ToString(),
                //               eci.BusinessName,
                //               eci.BusinessType
                //           };



                //var data2 = from a in context.Leaders
                //            join b in context.FamilyMembers
                //            on a.Id equals b.LeaderId into dc
                //            from dci in dc.DefaultIfEmpty()
                //            select new
                //            {
                //                LeaderId = a.Id,
                //                LeaderFullName = a.FullName,
                //                LeaderCardNo = a.CardNo,
                //                FamilyId = dci.Id.ToString(),
                //                dci.Relation,
                //                FamilyFullName = dci.FullName,
                //                FamilyCardNo = dci.CardNo
                //            };

                //var data3 = from a in context.WorkUnits
                //            join b in context.Leaders
                //            on a.Id equals b.WorkUnitId
                //            join c in context.FamilyMembers
                //            on b.Id equals c.LeaderId
                //            where (a.IsEntrust == false)
                //            select new
                //            {
                //                LeaderId = c.LeaderId,
                //                FullName = c.FullName,
                //                CardNo = c.CardNo,
                //                ReportDate = DateTime.Now
                //            };

                #endregion

                var workUnits = context.WorkUnits.Include(a => a.Leaders).ThenInclude(b => b.FamilyMembers).Where(a => a.IsEntrust == false).ToList();
                List<FamilyMember> familyMembers = new List<FamilyMember>();
                foreach (WorkUnit workUnit in workUnits)
                {
                    var leaders = workUnit.Leaders.OrderBy(b => b.OrderId).ToList();
                    foreach (Leader leader in leaders)
                    {
                        familyMembers.AddRange(leader.FamilyMembers.Where(a=>a.IsValid).OrderBy(b=>b.OrderId).ToList());
                    }
                    workUnit.IsEntrust = true;
                }
                CollectToFile(familyMembers);
                context.UpdateRange(workUnits);
                context.SaveChanges();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.dataGridView1.AutoGenerateColumns = false;

            using (BusinessContext context = new BusinessContext())
            {
                var workUnits = context.WorkUnits.ToList();
                
                if (workUnits != null)
                {
                    AutoCompleteStringCollection ac = new AutoCompleteStringCollection();

                    for (int i = 0; i < workUnits.Count; i++)
                    {
                        ac.Add(workUnits[i].UnitName);
                    }
                    this.listBox1.DataSource = ac;
                    this.textBox1.AutoCompleteCustomSource = ac;//设置该文本框的自动完成数据源

                    var data = context.Leaders.Where(b => b.WorkUnitId == workUnits[0].Id).ToList();
                    dataGridView1.DataSource = data;
                }                
            }
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (BusinessContext context = new BusinessContext())
            {
                var data = context.Leaders.Where(b => b.WorkUnit.UnitName == listBox1.SelectedValue.ToString()).ToList();
                dataGridView1.DataSource = data;
            }
                
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            using (BusinessContext context = new BusinessContext())
            {
                var data = context.FamilyMembers.Where(b => !b.IsValid).ToList();
                foreach(FamilyMember family in data)
                {
                    if (IDCardValidation.CheckIDCard(family.CardNo))
                    {
                        family.IsValid = true;
                        context.Update(family);
                    }

                }
                context.SaveChanges();
                var mCount = context.FamilyMembers.Where(b => !b.IsValid).Count();
                MessageBox.Show("共有" + mCount.ToString() + "位同志未通过身份证校验。");                
            }                
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            //List<WorkUnit> workUnits = new List<WorkUnit>();
            using(BusinessContext context = new BusinessContext())
            {
                var workUnits = context.WorkUnits.Include(a => a.Leaders).ThenInclude(b => b.FamilyMembers).ThenInclude(c=>c.ReportBusinesses).Where(a => a.IsEntrust == false).ToList();
                foreach(WorkUnit unit in workUnits)
                {
                    foreach(Leader leader in unit.Leaders)
                    {
                        CompareToFile(leader);
                    }
                }
                
            }


        }
        
    }
}
