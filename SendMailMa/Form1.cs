using ClosedXML.Excel;
using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using MimeKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using System.Linq;
using System.Data.Common.CommandTrees.ExpressionBuilder;
using System.Collections.Generic;
using SendMailMa.DATA;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Security.Policy;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;

namespace SendMailMa
{
    public partial class Form1 : Form
    {
        private readonly DATA.HelpdeskEntities db = new DATA.HelpdeskEntities();
        public Form1()
        {
            InitializeComponent();
        }


        private async void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "กำลังส่ง Email...";
            label1.ForeColor = System.Drawing.Color.Blue;
            label1.Font = new System.Drawing.Font(label1.Font, FontStyle.Bold);

            try
            {
                var db = new DATA.HelpdeskEntities();
                //DateTime today = Convert.ToDateTime("2026-03-01"); //Test 

                //------------------------------------------------------------------------------------------------------------------
                //ส่งเมล์ ล่วงหน้า 1 เดือน เช่น หมดอายุ(EndDate) ของเดือนที่ 2 ให้แจ้งเตือนตั้งแต่เดือนที่ 1 โดยส่งวันที่ 1 ของทุกเดือน
               DateTime today = DateTime.Today;

                // ส่งเมล์ล่วงหน้า 1 เดือน โดยส่งวันที่ 1 ของทุกเดือน
                if (today.Day == 1)
                //if (true)
                {
                    var mTargetDate = today.AddMonths(1);       // ถ้าวันนี้ 1 ม.ค. > ได้ 1 ก.พ.
                    int mTargetYear = mTargetDate.Year;        // 2025
                    int mTargetMonth = mTargetDate.Month;      // 2

                    string shortMonthName = mTargetDate.ToString("MMM", CultureInfo.CreateSpecificCulture("en-US"));



                    DateTime mStartDate = new DateTime(mTargetYear, mTargetMonth, 1); // 1 ก.พ.
                    DateTime mEndDate = new DateTime(mTargetYear, mTargetMonth, DateTime.DaysInMonth(mTargetYear, mTargetMonth));

                    var aeNamesMonth = db.VW_GetMA
                        .Where(x => (x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null) &&
                                    x.TRD_END_DATE >= mStartDate &&
                                    x.TRD_END_DATE <= mEndDate)
                        .Select(x => x.TRH_AE_NAME)
                        .Where(name => !string.IsNullOrEmpty(name))
                        .Distinct()
                        .ToList();

                    var result = new List<VW_GetMA>();
                    string url = "http://support.penso.co.th:84/";

                    foreach (var AeName in aeNamesMonth)
                    {
                        var maAeName = AeName;
                        var searchAeName = (AeName == "P Enterprise" || AeName == "CC Thai")
                        ? "Kanlayanee" //ที่ส่งไปหาพี่นุ้ยเพราะ บ. ไม่มี Email
                        : AeName;

                        result = db.VW_GetMA
                             .Where(x => (x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null) &&
                                         x.TRD_END_DATE >= mStartDate &&
                                         x.TRD_END_DATE <= mEndDate &&
                                         x.TRH_AE_NAME == maAeName)
                             .ToList();

                        var employee = db.VW_GetEmployeeReport.FirstOrDefault(x => x.name_eng == searchAeName);
                        if (employee == null || string.IsNullOrWhiteSpace(employee.emp_email))
                        {
                            Console.WriteLine("ไม่พบอีเมลของ AE: " + searchAeName);
                            return;
                        }

                        string GetEmail = employee.emp_email;

                        var GetEmailCC = (from x in db.VW_GetMailCC where x.EmpMail == GetEmail select x.EmpMailCC).ToList();

                        // สร้าง Excel
                        var workbook = new XLWorkbook();
                        var worksheet = workbook.Worksheets.Add("MA");

                        var titleRange = worksheet.Range(1, 1, 1, 18); // A1:R1 (18 คอลัมน์)
                        titleRange.Merge();
                        titleRange.Value = "MA Tracking(ล่วงหน้า 1 เดือน)";
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        titleRange.Style.Fill.BackgroundColor = XLColor.Green;     // พื้นหลังเขียว
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 16;
                        titleRange.Style.Font.FontColor = XLColor.White;

                        // Merge row 2 (A2:R2) เป็นแถวว่าง ไม่มีเส้น
                        var emptyRow = worksheet.Range(2, 1, 2, 18);
                        emptyRow.Merge();
                        emptyRow.Value = ""; // ไม่ต้องมีข้อความ
                        emptyRow.Style.Fill.BackgroundColor = XLColor.NoColor;
                        emptyRow.Style.Border.OutsideBorder = XLBorderStyleValues.None;
                        emptyRow.Style.Border.InsideBorder = XLBorderStyleValues.None;

                        // ---------- Header row at row 3 ----------
                        int headerRow = 3;
                        worksheet.Cell(headerRow, 1).Value = "NO.";
                        worksheet.Cell(headerRow, 2).Value = "Customer Name";
                        worksheet.Cell(headerRow, 3).Value = "End User";
                        worksheet.Cell(headerRow, 4).Value = "Charge Code";
                        worksheet.Cell(headerRow, 5).Value = "Project Name";
                        worksheet.Cell(headerRow, 6).Value = "AE";
                        worksheet.Cell(headerRow, 7).Value = "PO";
                        worksheet.Cell(headerRow, 8).Value = "Vendor";
                        worksheet.Cell(headerRow, 9).Value = "Brand";
                        worksheet.Cell(headerRow, 10).Value = "Code";
                        worksheet.Cell(headerRow, 11).Value = "Product Item";
                        worksheet.Cell(headerRow, 12).Value = "Qty";
                        worksheet.Cell(headerRow, 13).Value = "SN";
                        worksheet.Cell(headerRow, 14).Value = "Warranty";
                        worksheet.Cell(headerRow, 15).Value = "War";
                        worksheet.Cell(headerRow, 16).Value = "Start Date";
                        worksheet.Cell(headerRow, 17).Value = "End Date";
                        worksheet.Cell(headerRow, 18).Value = "Status";

                        //สไตล์หัวตารางเล็กน้อย(ถ้าต้องการ)
                        var headerRange = worksheet.Range(headerRow, 1, headerRow, 18);
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        int row = headerRow + 1; // = 4
                        int RowNumber = 1;
                        foreach (var item in result)
                        {
                            worksheet.Cell(row, 1).Value = RowNumber++;
                            worksheet.Cell(row, 2).Value = item.TRH_CUS_NAME;
                            worksheet.Cell(row, 3).Value = item.TRH_END_USER;
                            worksheet.Cell(row, 4).Value = item.TRH_PROJECT_CODE;
                            worksheet.Cell(row, 5).Value = item.TRH_PROJECT_NAME;
                            worksheet.Cell(row, 6).Value = item.TRH_AE_NAME;
                            worksheet.Cell(row, 7).Value = item.TRH_PO_NO;
                            worksheet.Cell(row, 8).Value = item.TRH_VEN_NAME;
                            worksheet.Cell(row, 9).Value = item.TRD_BRAND;
                            worksheet.Cell(row, 10).Value = item.TRD_CODE;
                            worksheet.Cell(row, 11).Value = item.TRD_PRODUCT_ITEM;
                            worksheet.Cell(row, 12).Value = item.TRD_QTY;
                            worksheet.Cell(row, 13).Value = item.TRSN_CODE;
                            worksheet.Cell(row, 14).Value = item.TRD_WARRANTY;
                            worksheet.Cell(row, 15).Value = item.TRD_WAR;
                            worksheet.Cell(row, 16).Value = item.TRD_START_DATE;
                            worksheet.Cell(row, 17).Value = item.TRD_END_DATE;
                            worksheet.Cell(row, 18).Value = item.TRD_STATUS;

                            worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 18).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            row++;
                        }

                        // AutoFit & Date Format
                        worksheet.Column(16).Style.DateFormat.Format = "yyyy-MM-dd"; // Start Date
                        worksheet.Column(17).Style.DateFormat.Format = "yyyy-MM-dd"; // End Date
                        worksheet.Columns().AdjustToContents();

                        // กันคอลัมน์แคบเกินไป
                        worksheet.Column(16).Width = Math.Max(worksheet.Column(16).Width, 12);
                        worksheet.Column(17).Width = Math.Max(worksheet.Column(17).Width, 12);

                        // ปรับขนาดอัตโนมัติ
                        worksheet.Columns().AdjustToContents();

                        // กรอบ cell
                        var dataRange = worksheet.Range($"A3:R{row - 1}");
                        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        // แปลงเป็น MemoryStream
                        var stream = new MemoryStream();
                        workbook.SaveAs(stream);
                        stream.Position = 0;

                        // เตรียม email
                        var message = new MimeMessage();
                        //message.From.Add(MailboxAddress.Parse("serviceccthailand@gmail.com"));
                        message.From.Add(MailboxAddress.Parse("matrackingsystemccpbg@gmail.com"));

                        //Test
                        //message.To.Add(MailboxAddress.Parse("theerawat@ccthailand.co.th"));
                        //message.To.Add(MailboxAddress.Parse("khaimook@penso.co.th"));
                        //message.To.Add(MailboxAddress.Parse("kanlayanee@penso.co.th"));
                        //message.To.Add(MailboxAddress.Parse("Ploypailin@penso.co.th"));

                        //Production
                        message.To.Add(MailboxAddress.Parse(GetEmail));
                        foreach (var item in GetEmailCC)
                        {
                            message.To.Add(MailboxAddress.Parse(item));
                        }

                        message.Subject = "MA TRACKING";

                        var builder = new BodyBuilder
                        {
                            TextBody = $"เรียนคุณ {searchAeName}\nไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel\nกรุณาตรวจสอบข้อมูลในระบบ\nหมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!"
                        };

                        builder.HtmlBody = $@"
                                            <p>เรียนคุณ {searchAeName}</p>
                                            <p>ไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel</p>
                                            <p>กรุณาตรวจสอบข้อมูลในระบบ 
                                               <a href=""{url}"">{url}</a>
                                            </p>
                                            <p><strong style=""color:#d32f2f;"">หมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!</strong></p>";

                        builder.Attachments.Add("MA_End_'" + shortMonthName + "'_'" + mTargetYear + "'.xlsx", stream.ToArray(),
                            new ContentType("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                        message.Body = builder.ToMessageBody();

                        //ส่งเมล
                        try
                        {
                            if (result.Count != 0)
                            {
                                var smtp = new SmtpClient();
                                await smtp.ConnectAsync("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                                await smtp.AuthenticateAsync("matrackingsystemccpbg@gmail.com", "vlkr knob dycv xunn");
                                await smtp.SendAsync(message);
                                await smtp.DisconnectAsync(true);

                                Console.WriteLine("ส่งอีเมลสำเร็จ");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ส่งอีเมลไม่สำเร็จ: " + ex.Message);
                        }
                    }
                }
                //------------------------------------------------------------------------------------------------------------------


                //------------------------------------------------------------------------------------------------------------------
                //ส่งเมล์ ล่วงหน้า 2 เดือน แบบ ส่งเป็น Quarter
                // ส่งเมลแบบ Q ล่วงหน้า 2 เดือน (รันทุกวัน แต่ทำงานเฉพาะวันที่ 1 ของเดือนที่กำหนด)
                if (today.Day == 1)
                //if (true)
                {
                    string Q = "";
                    int qStartMonth = 0;
                    int qYear = today.Year;

                    if (today.Month == 2)                     // 1 ก.พ.  Q2 (เม.ย.–มิ.ย.)
                    {
                        qStartMonth = 4;                      // เม.ย.
                        qYear = today.Year;                   // ปีเดียวกัน
                        Q = "Q2";
                    }
                    else if (today.Month == 5)                // 1 พ.ค.  Q3 (ก.ค.–ก.ย.)
                    {
                        qStartMonth = 7;                      // ก.ค.
                        qYear = today.Year;
                        Q = "Q3";
                    }
                    else if (today.Month == 8)                // 1 ส.ค.  Q4 (ต.ค.–ธ.ค.)
                    {
                        qStartMonth = 10;                     // ต.ค.
                        qYear = today.Year;
                        Q = "Q4";
                    }
                    else if (today.Month == 11)               // 1 พ.ย.  Q1 (ม.ค.–มี.ค. ปีหน้า)
                    {
                        qStartMonth = 1;                      // ม.ค.
                        qYear = today.Year + 1;               // ปีถัดไป
                        Q = "Q1";
                    }
                    else
                    {
                        // ถ้าเป็นเดือนอื่น ๆ ไม่ส่งแบบ Q
                        qStartMonth = 0;
                    }

                    if (qStartMonth != 0)
                    {

                        DateTime qStartDate = new DateTime(qYear, qStartMonth, 1);
                        DateTime qEndDateExclusive = qStartDate.AddMonths(3);

                        var aeNamesQuarter = db.VW_GetMA
                   .Where(x => (x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null) &&
                               x.TRD_END_DATE >= qStartDate &&
                               x.TRD_END_DATE <= qEndDateExclusive)
                   .Select(x => x.TRH_AE_NAME)
                   .Where(name => !string.IsNullOrEmpty(name))
                   .Distinct()
                   .ToList();


                        var result = new List<VW_GetMA>();

                        foreach (var AeName in aeNamesQuarter)
                        {
                            var maAeName = AeName;
                            var searchAeName = (AeName == "P Enterprise" || AeName == "CC Thai")
                            ? "Kanlayanee" //ที่ส่งไปหาพี่นุ้ยเพราะ บ. ไม่มี Email
                            : AeName;

                            result = db.VW_GetMA
                                .Where(x => (x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null)  &&
                                            x.TRD_END_DATE >= qStartDate &&
                                            x.TRD_END_DATE < qEndDateExclusive &&
                                            x.TRH_AE_NAME == maAeName)
                                .ToList();

                         
                            string url = "http://support.penso.co.th:84/";
                            var employee = db.VW_GetEmployeeReport.FirstOrDefault(x => x.name_eng == searchAeName);
                        
                            if (employee == null || string.IsNullOrWhiteSpace(employee.emp_email))
                            {
                                Console.WriteLine("ไม่พบอีเมลของ AE: " + searchAeName);
                                return;
                            }

                            //Select Email ที่ต้องการส่ง จากเอกสาร
                            string GetEmail = employee.emp_email;

                            var GetEmailCC = (from x in db.VW_GetMailCC where x.EmpMail == GetEmail select x.EmpMailCC).ToList();

                            // สร้าง Excel
                            var workbook = new XLWorkbook();
                            var worksheet = workbook.Worksheets.Add("MA");

                            var titleRange = worksheet.Range(1, 1, 1, 18); // A1:R1 (18 คอลัมน์)
                            titleRange.Merge();
                            titleRange.Value = "MA Tracking(ล่วงหน้า 2 เดือนแบบ Quarter)";
                            titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            titleRange.Style.Fill.BackgroundColor = XLColor.Green;     // พื้นหลังเขียว
                            titleRange.Style.Font.Bold = true;
                            titleRange.Style.Font.FontSize = 16;
                            titleRange.Style.Font.FontColor = XLColor.White;

                            // Merge row 2 (A2:R2) เป็นแถวว่าง ไม่มีเส้น
                            var emptyRow = worksheet.Range(2, 1, 2, 18);
                            emptyRow.Merge();
                            emptyRow.Value = ""; // ไม่ต้องมีข้อความ
                            emptyRow.Style.Fill.BackgroundColor = XLColor.NoColor;
                            emptyRow.Style.Border.OutsideBorder = XLBorderStyleValues.None;
                            emptyRow.Style.Border.InsideBorder = XLBorderStyleValues.None;

                            // ---------- Header row at row 3 ----------
                            int headerRow = 3;
                            worksheet.Cell(headerRow, 1).Value = "NO.";
                            worksheet.Cell(headerRow, 2).Value = "Customer Name";
                            worksheet.Cell(headerRow, 3).Value = "End User";
                            worksheet.Cell(headerRow, 4).Value = "Charge Code";
                            worksheet.Cell(headerRow, 5).Value = "Project Name";
                            worksheet.Cell(headerRow, 6).Value = "AE";
                            worksheet.Cell(headerRow, 7).Value = "PO";
                            worksheet.Cell(headerRow, 8).Value = "Vendor";
                            worksheet.Cell(headerRow, 9).Value = "Brand";
                            worksheet.Cell(headerRow, 10).Value = "Code";
                            worksheet.Cell(headerRow, 11).Value = "Product Item";
                            worksheet.Cell(headerRow, 12).Value = "Qty";
                            worksheet.Cell(headerRow, 13).Value = "SN";
                            worksheet.Cell(headerRow, 14).Value = "Warranty";
                            worksheet.Cell(headerRow, 15).Value = "War";
                            worksheet.Cell(headerRow, 16).Value = "Start Date";
                            worksheet.Cell(headerRow, 17).Value = "End Date";
                            worksheet.Cell(headerRow, 18).Value = "Status";

                            //สไตล์หัวตารางเล็กน้อย(ถ้าต้องการ)
                            var headerRange = worksheet.Range(headerRow, 1, headerRow, 18);
                            headerRange.Style.Font.Bold = true;
                            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                            int row = headerRow + 1; // = 4
                            int RowNumber = 1;
                            foreach (var item in result)
                            {
                                worksheet.Cell(row, 1).Value = RowNumber++;
                                worksheet.Cell(row, 2).Value = item.TRH_CUS_NAME;
                                worksheet.Cell(row, 3).Value = item.TRH_END_USER;
                                worksheet.Cell(row, 4).Value = item.TRH_PROJECT_CODE;
                                worksheet.Cell(row, 5).Value = item.TRH_PROJECT_NAME;
                                worksheet.Cell(row, 6).Value = item.TRH_AE_NAME;
                                worksheet.Cell(row, 7).Value = item.TRH_PO_NO;
                                worksheet.Cell(row, 8).Value = item.TRH_VEN_NAME;
                                worksheet.Cell(row, 9).Value = item.TRD_BRAND;
                                worksheet.Cell(row, 10).Value = item.TRD_CODE;
                                worksheet.Cell(row, 11).Value = item.TRD_PRODUCT_ITEM;
                                worksheet.Cell(row, 12).Value = item.TRD_QTY;
                                worksheet.Cell(row, 13).Value = item.TRSN_CODE;
                                worksheet.Cell(row, 14).Value = item.TRD_WARRANTY;
                                worksheet.Cell(row, 15).Value = item.TRD_WAR;
                                worksheet.Cell(row, 16).Value = item.TRD_START_DATE;
                                worksheet.Cell(row, 17).Value = item.TRD_END_DATE;
                                worksheet.Cell(row, 18).Value = item.TRD_STATUS;

                                worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                worksheet.Cell(row, 18).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                                row++;
                            }

                            // AutoFit & Date Format
                            worksheet.Column(16).Style.DateFormat.Format = "yyyy-MM-dd"; // Start Date
                            worksheet.Column(17).Style.DateFormat.Format = "yyyy-MM-dd"; // End Date
                            worksheet.Columns().AdjustToContents();

                            // กันคอลัมน์แคบเกินไป
                            worksheet.Column(16).Width = Math.Max(worksheet.Column(16).Width, 12);
                            worksheet.Column(17).Width = Math.Max(worksheet.Column(17).Width, 12);

                            // ปรับขนาดอัตโนมัติ
                            worksheet.Columns().AdjustToContents();

                            // กรอบ cell
                            var dataRange = worksheet.Range($"A3:R{row - 1}");
                            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                            // แปลงเป็น MemoryStream
                            var stream = new MemoryStream();
                            workbook.SaveAs(stream);
                            stream.Position = 0;

                            // เตรียม email
                            var message = new MimeMessage();
                            message.From.Add(MailboxAddress.Parse("matrackingsystemccpbg@gmail.com"));

                            //Test
                            //message.To.Add(MailboxAddress.Parse("theerawat@ccthailand.co.th"));
                            //message.To.Add(MailboxAddress.Parse("khaimook@penso.co.th"));
                            //message.To.Add(MailboxAddress.Parse("kanlayanee@penso.co.th"));
                            //message.To.Add(MailboxAddress.Parse("Ploypailin@penso.co.th"));

                            //Production
                            message.To.Add(MailboxAddress.Parse(GetEmail));
                            foreach (var item in GetEmailCC)
                            {
                                message.To.Add(MailboxAddress.Parse(item));
                            }
                            message.Subject = "MA TRACKING";

                            var builder = new BodyBuilder
                            {
                                TextBody = $"เรียนคุณ {searchAeName}\nไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel\nกรุณาตรวจสอบข้อมูลในระบบ\nหมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!"
                            };

                            builder.HtmlBody = $@"
                                            <p>เรียนคุณ {searchAeName}</p>
                                            <p>ไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel</p>
                                            <p>กรุณาตรวจสอบข้อมูลในระบบ 
                                               <a href=""{url}"">{url}</a>
                                            </p>
                                            <p><strong style=""color:#d32f2f;"">หมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!</strong></p>";

                            builder.Attachments.Add("MA_END_" + Q + "_" + qYear + ".xlsx", stream.ToArray(),
                                new ContentType("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                            message.Body = builder.ToMessageBody();

                            //ส่งเมล
                            try
                            {
                                if (result.Count != 0)
                                {
                                    var smtp = new SmtpClient();
                                    await smtp.ConnectAsync("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                                    await smtp.AuthenticateAsync("matrackingsystemccpbg@gmail.com", "vlkr knob dycv xunn");
                                    await smtp.SendAsync(message);
                                    await smtp.DisconnectAsync(true);

                                    Console.WriteLine("ส่งอีเมลสำเร็จ");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("ส่งอีเมลไม่สำเร็จ: " + ex.Message);
                            }
                        }
                    }
                }
                //------------------------------------------------------------------------------------------------------------------


                //------------------------------------------------------------------------------------------------------------------
                //ส่งข้อมูลทุกวันที่ 15 ของเดือน เช่น วันนี้ 12 มกราคม ให้ส่งข้อมูล เดือนก่อนหน้าเดือนปัจจุบันลงไปทั้งหมด ที่สถานะเป็น Pending
                if (today.Day == 15)
                //if (true)
                {
                    string shortMonthName = today.ToString("MMM", CultureInfo.CreateSpecificCulture("en-US"));
                    int year = today.Year;
                    DateTime GetMonthEnd = new DateTime(today.Year, today.Month, 1).AddDays(-1);

                    var aeNamesMonth = db.VW_GetMA
                        .Where(x =>
                            (x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null) &&
                            x.TRD_END_DATE <= GetMonthEnd
                        )
                        .Select(x => x.TRH_AE_NAME)
                        .Where(name => !string.IsNullOrEmpty(name))
                        .Distinct()
                        .ToList();

                    var result = new List<VW_GetMA>();
                    string url = "http://support.penso.co.th:84/";

                    foreach (var AeName in aeNamesMonth)
                    {
                        var maAeName = AeName;
                        var searchAeName = (AeName == "P Enterprise" || AeName == "CC Thai")
                        ? "Kanlayanee" //ที่ส่งไปหาพี่นุ้ยเพราะ บ. ไม่มี Email
                        : AeName;

                        result = db.VW_GetMA
                            .Where(x =>
                                (x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null) &&
                                x.TRD_END_DATE <= GetMonthEnd &&
                                x.TRH_AE_NAME == maAeName
                            )
                            .ToList();


                        var employee = db.VW_GetEmployeeReport.FirstOrDefault(x => x.name_eng == searchAeName);

                        if (employee == null || string.IsNullOrWhiteSpace(employee.emp_email))
                        {
                            Console.WriteLine("ไม่พบอีเมลของ AE: " + searchAeName);
                            return;
                        }

                        // สร้าง Excel
                        var workbook = new XLWorkbook();
                        var worksheet = workbook.Worksheets.Add("MA");

                        var titleRange = worksheet.Range(1, 1, 1, 18); // A1:R1 (18 คอลัมน์)
                        titleRange.Merge();
                        titleRange.Value = "MA Tracking(Recheck 15 วัน)";
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        titleRange.Style.Fill.BackgroundColor = XLColor.Green;     // พื้นหลังเขียว
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 16;
                        titleRange.Style.Font.FontColor = XLColor.White;

                        // Merge row 2 (A2:R2) เป็นแถวว่าง ไม่มีเส้น
                        var emptyRow = worksheet.Range(2, 1, 2, 18);
                        emptyRow.Merge();
                        emptyRow.Value = ""; // ไม่ต้องมีข้อความ
                        emptyRow.Style.Fill.BackgroundColor = XLColor.NoColor;
                        emptyRow.Style.Border.OutsideBorder = XLBorderStyleValues.None;
                        emptyRow.Style.Border.InsideBorder = XLBorderStyleValues.None;

                        // ---------- Header row at row 3 ----------
                        int headerRow = 3;
                        worksheet.Cell(headerRow, 1).Value = "NO.";
                        worksheet.Cell(headerRow, 2).Value = "Customer Name";
                        worksheet.Cell(headerRow, 3).Value = "End User";
                        worksheet.Cell(headerRow, 4).Value = "Charge Code";
                        worksheet.Cell(headerRow, 5).Value = "Project Name";
                        worksheet.Cell(headerRow, 6).Value = "AE";
                        worksheet.Cell(headerRow, 7).Value = "PO";
                        worksheet.Cell(headerRow, 8).Value = "Vendor";
                        worksheet.Cell(headerRow, 9).Value = "Brand";
                        worksheet.Cell(headerRow, 10).Value = "Code";
                        worksheet.Cell(headerRow, 11).Value = "Product Item";
                        worksheet.Cell(headerRow, 12).Value = "Qty";
                        worksheet.Cell(headerRow, 13).Value = "SN";
                        worksheet.Cell(headerRow, 14).Value = "Warranty";
                        worksheet.Cell(headerRow, 15).Value = "War";
                        worksheet.Cell(headerRow, 16).Value = "Start Date";
                        worksheet.Cell(headerRow, 17).Value = "End Date";
                        worksheet.Cell(headerRow, 18).Value = "Status";

                        //สไตล์หัวตารางเล็กน้อย(ถ้าต้องการ)
                        var headerRange = worksheet.Range(headerRow, 1, headerRow, 18);
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        int row = headerRow + 1; // = 4
                        int RowNumber = 1;
                        foreach (var item in result)
                        {
                            worksheet.Cell(row, 1).Value = RowNumber++;
                            worksheet.Cell(row, 2).Value = item.TRH_CUS_NAME;
                            worksheet.Cell(row, 3).Value = item.TRH_END_USER;
                            worksheet.Cell(row, 4).Value = item.TRH_PROJECT_CODE;
                            worksheet.Cell(row, 5).Value = item.TRH_PROJECT_NAME;
                            worksheet.Cell(row, 6).Value = item.TRH_AE_NAME;
                            worksheet.Cell(row, 7).Value = item.TRH_PO_NO;
                            worksheet.Cell(row, 8).Value = item.TRH_VEN_NAME;
                            worksheet.Cell(row, 9).Value = item.TRD_BRAND;
                            worksheet.Cell(row, 10).Value = item.TRD_CODE;
                            worksheet.Cell(row, 11).Value = item.TRD_PRODUCT_ITEM;
                            worksheet.Cell(row, 12).Value = item.TRD_QTY;
                            worksheet.Cell(row, 13).Value = item.TRSN_CODE;
                            worksheet.Cell(row, 14).Value = item.TRD_WARRANTY;
                            worksheet.Cell(row, 15).Value = item.TRD_WAR;
                            worksheet.Cell(row, 16).Value = item.TRD_START_DATE;
                            worksheet.Cell(row, 17).Value = item.TRD_END_DATE;
                            worksheet.Cell(row, 18).Value = item.TRD_STATUS;

                            worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 18).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            row++;
                        }

                        // AutoFit & Date Format
                        worksheet.Column(16).Style.DateFormat.Format = "yyyy-MM-dd"; // Start Date
                        worksheet.Column(17).Style.DateFormat.Format = "yyyy-MM-dd"; // End Date
                        worksheet.Columns().AdjustToContents();

                        // กันคอลัมน์แคบเกินไป
                        worksheet.Column(16).Width = Math.Max(worksheet.Column(16).Width, 12);
                        worksheet.Column(17).Width = Math.Max(worksheet.Column(17).Width, 12);

                        // ปรับขนาดอัตโนมัติ
                        worksheet.Columns().AdjustToContents();

                        // กรอบ cell
                        var dataRange = worksheet.Range($"A3:R{row - 1}");
                        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        // แปลงเป็น MemoryStream
                        var stream = new MemoryStream();
                        workbook.SaveAs(stream);
                        stream.Position = 0;

                        string GetEmail = employee.emp_email;

                        var GetEmailCC = (from x in db.VW_GetMailCC where x.EmpMail == GetEmail select x.EmpMailCC).ToList();

                        // เตรียม email
                        var message = new MimeMessage();
                        message.From.Add(MailboxAddress.Parse("matrackingsystemccpbg@gmail.com"));

                        //Test
                        //message.To.Add(MailboxAddress.Parse("theerawat@ccthailand.co.th"));


                        //Production
                        message.To.Add(MailboxAddress.Parse(GetEmail));
                        foreach (var item in GetEmailCC)
                        {
                            message.To.Add(MailboxAddress.Parse(item));
                        }
                        message.Subject = "MA TRACKING";

                        var builder = new BodyBuilder
                        {
                            TextBody = $"เรียนคุณ {searchAeName}\nไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel\nกรุณาตรวจสอบข้อมูลในระบบ\nหมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!"
                        };

                        builder.HtmlBody = $@"
                                            <p>เรียนคุณ {searchAeName}</p>
                                            <p>ไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel</p>
                                            <p>กรุณาตรวจสอบข้อมูลในระบบ 
                                               <a href=""{url}"">{url}</a>
                                            </p>
                                            <p><strong style=""color:#d32f2f;"">หมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!</strong></p>";

                        builder.Attachments.Add("MA_End_Before_" + shortMonthName + "_" + year + ".xlsx", stream.ToArray(),
                            new ContentType("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                        message.Body = builder.ToMessageBody();

                        // ส่งเมล
                        try
                        {
                            if (result.Count != 0)
                            {
                                var smtp = new SmtpClient();
                                await smtp.ConnectAsync("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                                await smtp.AuthenticateAsync("matrackingsystemccpbg@gmail.com", "vlkr knob dycv xunn");
                                await smtp.SendAsync(message);
                                await smtp.DisconnectAsync(true);

                                Console.WriteLine("ส่งอีเมลสำเร็จ");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ส่งอีเมลไม่สำเร็จ: " + ex.Message);
                        }
                    }
                }


                //------------------------------------------------------------------------------------------------------------------
                // ส่งข้อมูลทุกวันที่ 15 ของเดือน Status Pending ส่งข้อมูลเดือนนก่อนหน้าลงไปทั้งหมด Status Not-renewed ส่งข้อมูลเดือนก่อนหน้า 1 เดือน
                if (today.Day == 15)
                //if (true)
                {
                    DateTime GetMonthEnd = new DateTime(today.Year, today.Month, 1).AddDays(-1);
                    // ต้นเดือน–ต้นเดือนถัดไป
                    var monthStart = new DateTime(GetMonthEnd.Year, GetMonthEnd.Month, 1);
                    var nextMonthStart = monthStart.AddMonths(1);

                    string shortMonthName = today.ToString("MMM", CultureInfo.CreateSpecificCulture("en-US"));
                    int year = today.Year;

                    var aeNamesMonth = db.VW_GetMA
                        .Where(x =>
                            ((x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null) ||
                            x.TRD_STATUS_VALUE == "0") &&
                            x.TRD_END_DATE <= GetMonthEnd
                        )
                        .Select(x => x.TRH_AE_NAME)
                        .Where(name => !string.IsNullOrEmpty(name))
                        .Distinct()
                        .ToList();

                    var resultPending = new List<VW_GetMA>();
                    var resultNotRenewed = new List<VW_GetMA>();
                    string url = "http://support.penso.co.th:84/";
                    foreach (var AeName in aeNamesMonth)
                    {
                        var maAeName = AeName;
                        var searchAeName = (AeName == "P Enterprise" || AeName == "CC Thai")
                        ? "Kanlayanee" //ที่ส่งไปหาพี่นุ้ยเพราะ บ. ไม่มี Email
                        : AeName;

                        // ดึงข้อมูล
                        resultPending = db.VW_GetMA
                            .Where(x =>
                                (x.TRD_STATUS_VALUE == "" || x.TRD_STATUS_VALUE == null) &&
                                x.TRD_END_DATE <= GetMonthEnd &&
                                x.TRH_AE_NAME == maAeName
                            )
                            .ToList();

                        resultNotRenewed = db.VW_GetMA
                            .Where(x =>
                                x.TRD_STATUS_VALUE == "0" &&
                                  x.TRD_END_DATE >= monthStart &&
                                  x.TRD_END_DATE < nextMonthStart &&
                                x.TRH_AE_NAME == maAeName
                            )
                            .ToList();

                        // ถ้าไม่มีทั้ง Pending และ Not-Renewed → ไม่ต้องสร้างเมล
                        if (resultPending.Count == 0 && resultNotRenewed.Count == 0)
                        {
                            continue;
                        }

                        // เลือก AE (ตอนนี้ fix เป็น emp_id = 9)
                        string AEName = db.VW_GetEmployeeReport
                                        .Where(x => x.emp_id == 9)   // ID พี่ต่อ
                                        .Select(x => x.name_eng)
                                        .Where(name => !string.IsNullOrEmpty(name))
                                        .Distinct()
                                        .SingleOrDefault();

                        var employee = db.VW_GetEmployeeReport.FirstOrDefault(x => x.name_eng == AEName);
                        if (employee == null || string.IsNullOrWhiteSpace(employee.emp_email))
                        {
                            Console.WriteLine("ไม่พบอีเมลของ AE: " + AEName);
                            continue;   // หรือ return; แล้วแต่ logic ที่ต้องการ
                        }

                        string GetEmail = employee.emp_email;
                        var GetEmailCC = (from x in db.VW_GetMailCC
                                          where x.EmpMail == GetEmail
                                          select x.EmpMailCC).ToList();

                        // เตรียม email
                        var message = new MimeMessage();
                        message.From.Add(MailboxAddress.Parse("matrackingsystemccpbg@gmail.com"));

                        // Test
                        //message.To.Add(MailboxAddress.Parse("theerawat@ccthailand.co.th"));
                        //message.To.Add(MailboxAddress.Parse("khaimook@penso.co.th"));
                        //message.To.Add(MailboxAddress.Parse("kanlayanee@penso.co.th"));
                        //message.To.Add(MailboxAddress.Parse("Ploypailin@penso.co.th"));

                        // Production (ปลดตอนใช้จริง)
                        message.To.Add(MailboxAddress.Parse(GetEmail));
                        foreach (var cc in GetEmailCC)
                        {
                            message.Cc.Add(MailboxAddress.Parse(cc));
                        }

                        message.Subject = "MA TRACKING";

                        var builder = new BodyBuilder
                        {
                            TextBody = $"เรียนคุณ {AEName}\nไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel\nกรุณาตรวจสอบข้อมูลในระบบ\nหมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!"
                        };

                        builder.HtmlBody = $@"
                                            <p>เรียนคุณ {AEName}</p>
                                            <p>ไฟล์แนบเอกสารครบกำหนดหมดอายุ (MA) ในรูปแบบ Excel</p>
                                            <p>กรุณาตรวจสอบข้อมูลในระบบ 
                                               <a href=""{url}"">{url}</a>
                                            </p>
                                            <p><strong style=""color:#d32f2f;"">หมายเหตุ: เป็น Email อัตโนมัติห้ามตอบกลับ!</strong></p>";

                        // แนบไฟล์เฉพาะที่มีข้อมูลจริง ๆ
                        if (resultPending.Count > 0)
                        {
                            var streamPending = CreateMaExcel(resultPending,
                                "MA Tracking(Recheck 15 วัน) - Pending");

                            builder.Attachments.Add(
                                $"MA_END_Before_{shortMonthName}_{year}(Pending).xlsx",
                                streamPending.ToArray(),
                                new ContentType("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                        }

                        if (resultNotRenewed.Count > 0)
                        {
                            var streamNotRenewed = CreateMaExcel(resultNotRenewed,
                                "MA Tracking(Recheck 15 วัน) - Not Renewed");

                            builder.Attachments.Add(
                                $"MA_END_Before_{shortMonthName}_{year}(Not-Renewed).xlsx",
                                streamNotRenewed.ToArray(),
                                new ContentType("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                        }

                        message.Body = builder.ToMessageBody();

                        try
                        {
                            var smtp = new SmtpClient();
                            await smtp.ConnectAsync("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                            await smtp.AuthenticateAsync("matrackingsystemccpbg@gmail.com", "vlkr knob dycv xunn");
                            await smtp.SendAsync(message);
                            await smtp.DisconnectAsync(true);

                            Console.WriteLine("ส่งอีเมลสำเร็จ");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ส่งอีเมลไม่สำเร็จ: " + ex.Message);
                        }
                    }



                }

                label1.Text = "ส่ง Email เรียบร้อย";
                label1.ForeColor = System.Drawing.Color.Green;
            }
            catch (Exception ex)
            {
                label1.Text = "ส่ง Email ล้มเหลว";
                label1.ForeColor = System.Drawing.Color.Red;
                MessageBox.Show("เกิดข้อผิดพลาด: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            await Task.Delay(5000);
            this.Close();
        }


        private MemoryStream CreateMaExcel(List<VW_GetMA> data, string title)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("MA");

            // Title
            var titleRange = worksheet.Range(1, 1, 1, 18); // A1:R1
            titleRange.Merge();
            titleRange.Value = title;
            titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            titleRange.Style.Fill.BackgroundColor = XLColor.Green;
            titleRange.Style.Font.Bold = true;
            titleRange.Style.Font.FontSize = 16;
            titleRange.Style.Font.FontColor = XLColor.White;

            // Row 2 ว่าง
            var emptyRow = worksheet.Range(2, 1, 2, 18);
            emptyRow.Merge();
            emptyRow.Value = "";
            emptyRow.Style.Fill.BackgroundColor = XLColor.NoColor;
            emptyRow.Style.Border.OutsideBorder = XLBorderStyleValues.None;
            emptyRow.Style.Border.InsideBorder = XLBorderStyleValues.None;

            // Header row = 3
            int headerRow = 3;
            worksheet.Cell(headerRow, 1).Value = "NO.";
            worksheet.Cell(headerRow, 2).Value = "Customer Name";
            worksheet.Cell(headerRow, 3).Value = "End User";
            worksheet.Cell(headerRow, 4).Value = "Charge Code";
            worksheet.Cell(headerRow, 5).Value = "Project Name";
            worksheet.Cell(headerRow, 6).Value = "AE";
            worksheet.Cell(headerRow, 7).Value = "PO";
            worksheet.Cell(headerRow, 8).Value = "Vendor";
            worksheet.Cell(headerRow, 9).Value = "Brand";
            worksheet.Cell(headerRow, 10).Value = "Code";
            worksheet.Cell(headerRow, 11).Value = "Product Item";
            worksheet.Cell(headerRow, 12).Value = "Qty";
            worksheet.Cell(headerRow, 13).Value = "SN";
            worksheet.Cell(headerRow, 14).Value = "Warranty";
            worksheet.Cell(headerRow, 15).Value = "War";
            worksheet.Cell(headerRow, 16).Value = "Start Date";
            worksheet.Cell(headerRow, 17).Value = "End Date";
            worksheet.Cell(headerRow, 18).Value = "Status";

            var headerRange = worksheet.Range(headerRow, 1, headerRow, 18);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Data
            int row = headerRow + 1; // เริ่มแถวที่ 4
            int no = 1;

            foreach (var item in data)
            {
                worksheet.Cell(row, 1).Value = no++;
                worksheet.Cell(row, 2).Value = item.TRH_CUS_NAME;
                worksheet.Cell(row, 3).Value = item.TRH_END_USER;
                worksheet.Cell(row, 4).Value = item.TRH_PROJECT_CODE;
                worksheet.Cell(row, 5).Value = item.TRH_PROJECT_NAME;
                worksheet.Cell(row, 6).Value = item.TRH_AE_NAME;
                worksheet.Cell(row, 7).Value = item.TRH_PO_NO;
                worksheet.Cell(row, 8).Value = item.TRH_VEN_NAME;
                worksheet.Cell(row, 9).Value = item.TRD_BRAND;
                worksheet.Cell(row, 10).Value = item.TRD_CODE;
                worksheet.Cell(row, 11).Value = item.TRD_PRODUCT_ITEM;
                worksheet.Cell(row, 12).Value = item.TRD_QTY;
                worksheet.Cell(row, 13).Value = item.TRSN_CODE;
                worksheet.Cell(row, 14).Value = item.TRD_WARRANTY;
                worksheet.Cell(row, 15).Value = item.TRD_WAR;
                worksheet.Cell(row, 16).Value = item.TRD_START_DATE;
                worksheet.Cell(row, 17).Value = item.TRD_END_DATE;
                worksheet.Cell(row, 18).Value = item.TRD_STATUS;

                // จัด alignment
                worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 18).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                row++;
            }

            // Date format & Autofit
            worksheet.Column(16).Style.DateFormat.Format = "yyyy-MM-dd";
            worksheet.Column(17).Style.DateFormat.Format = "yyyy-MM-dd";
            worksheet.Columns().AdjustToContents();
            worksheet.Column(16).Width = Math.Max(worksheet.Column(16).Width, 12);
            worksheet.Column(17).Width = Math.Max(worksheet.Column(17).Width, 12);
            worksheet.Columns().AdjustToContents();

            // กรอบ cell ของ header+data
            if (row - 1 >= headerRow)
            {
                var dataRange = worksheet.Range($"A3:R{row - 1}");
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }

            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            ms.Position = 0;
            return ms;
        }
    }
}