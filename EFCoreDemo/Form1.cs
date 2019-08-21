using System;
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
                    foreach (Leader leader in workUnit.Leaders)
                    {
                        familyMembers.AddRange(leader.FamilyMembers.OrderBy(b=>b.OrderId));
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
    }
}
