
using Aspose.Cells;
using Aspose.Cells.Charts;
using PugPdf.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace invoice
{
    class Program
    {
        public static string style = @"<style>
    table, td, th {
    border:1px solid;
}
* {
     font-size: 100%;
     font-family: Times New Roman, Times, serif;
}
 table {
    width: 100%;
     border-collapse: collapse;
     table-layout: auto;
}
 td {
    font - weight: bold;
     overflow-wrap: break-word;
     word-wrap: break-word;
}
 .center {
    margin: auto;
     width: 90%;
     padding: 10px;
}
 
  </style>";
        public static List<company> companyList = new List<company>();
        public static List<allStaff> allStaffList = new List<allStaff>();
        public static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public static string desktopFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\";
        public static string timeSheetFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\input\\";
        public static string outputFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\output\\";
        public static string outputStaffInvoiceFolder = string.Empty;
        public static string generalFolder = desktopPath + "\\" + "invoice\\general\\";
        public static string outputCompanyInvoiceHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\ouputHtml\\company\\";
        public static string outputCompanyInvoiceReceiptHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\ouputHtml\\company\\receipt\\";
        public static string outputStaffInvoiceHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\ouputHtml\\staff\\";
        public static string inputHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\inputHtml\\";
        public static string outputHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\inputHtml\\output\\";

        public static List<companyDuty> companyDuties = new List<companyDuty>();
        public static string invoiceDate = string.Empty;
        public static totalAmountObj companyTotalAmountList = new totalAmountObj();
        public static totalAmountObj staffTotalAmountList = new totalAmountObj();
        public static string check1 = string.Empty;
        static async Task Main(string[] args)
        {

            try
            {

                Console.OutputEncoding = Encoding.Unicode;




                if (!Directory.Exists(desktopFolder))
                {
                    Directory.CreateDirectory(desktopFolder);
                }
                if (!Directory.Exists(timeSheetFolder))
                {
                    Directory.CreateDirectory(timeSheetFolder);
                }
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                if (!Directory.Exists(generalFolder))
                {
                    Directory.CreateDirectory(generalFolder);
                }
                if (!Directory.Exists(outputFolder + "staff_invoice\\"))
                {
                    Directory.CreateDirectory(outputFolder + "staff_invoice\\");
                }
                if (!Directory.Exists(outputCompanyInvoiceHtmlFolder))
                {
                    Directory.CreateDirectory(outputCompanyInvoiceHtmlFolder);
                }
                if (!Directory.Exists(inputHtmlFolder))
                {
                    Directory.CreateDirectory(inputHtmlFolder);
                }
                if (!Directory.Exists(outputHtmlFolder))
                {
                    Directory.CreateDirectory(outputHtmlFolder);
                }
                if (!Directory.Exists(outputStaffInvoiceHtmlFolder))
                {
                    Directory.CreateDirectory(outputStaffInvoiceHtmlFolder);
                }
                if (!Directory.Exists(outputCompanyInvoiceReceiptHtmlFolder))
                {
                    Directory.CreateDirectory(outputCompanyInvoiceReceiptHtmlFolder);
                }


                string url = $"https://testsds123-669967cd5270.herokuapp.com/";
                using var client = new HttpClient();
                var response = client.GetAsync(url).GetAwaiter().GetResult();
                var content = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();// response value*/

                Console.WriteLine("將 timesheet.xlsx 放入資料夾 " + timeSheetFolder);
                Console.WriteLine("將 logo.jpg和company_salary.xlsx和staff_salary.xlsx和bank.xlsx 放入資料夾 " + generalFolder);
                Console.WriteLine("1 = Gen Company Invoice, 2 = Gen Staff Invoice, 3 = Gen HtmlCode");
                // string content = "success";
                check1 = Console.ReadLine();
                if (content == "success")
                {

                    if (check1 == "3")
                    {
                        var renderer = new HtmlToPdf();
                        DirectoryInfo d = new DirectoryInfo(inputHtmlFolder);
                        var matchFolder = d.GetFiles("*.txt");
                        for (int i = 0; i < matchFolder.Length; i++)
                        {
                            Console.WriteLine(matchFolder[i].FullName + " Processing");
                            string text = File.ReadAllText(matchFolder[i].FullName);
                            var pdf = await renderer.RenderHtmlAsPdfAsync(text);
                            pdf.SaveAs(outputHtmlFolder + Path.GetFileNameWithoutExtension(matchFolder[i].FullName) + ".pdf");
                            Console.WriteLine(matchFolder[i].FullName + " Done");
                        }

                    }
                    if (check1 == "1" || check1 == "2")
                    {



                        List<string> JobIdList = new List<string>();

                        //string content = "success";


                        Workbook wb = new Workbook(timeSheetFolder + "timesheet.xlsx");


                        for (int v = 0; v < wb.Worksheets.Count(); v++)
                        {
                            Console.WriteLine();
                            Console.WriteLine(v + "/" + wb.Worksheets.Count());
                            List<specialEvent> specialEventsList = new List<specialEvent>();
                            List<staffList> staffNameList = new List<staffList>();
                            company company = new company();

                            Worksheet worksheet = wb.Worksheets[v];
                            var ggg = worksheet.IsVisible;

                            if (worksheet.IsVisible == false)
                            {
                                Console.WriteLine("Hidden Table");
                                continue;
                            }

                            Console.WriteLine("Worksheet: " + worksheet.Name);

                            int rows = worksheet.Cells.MaxDataRow;
                            int cols = worksheet.Cells.MaxDataColumn;


                            List<string> timeList = new List<string>();
                            List<string> tempStaffNameList = new List<string>();


                            List<string> dateList = new List<string>();


                            company.customerName = worksheet.Cells[3, 4].Value != null ? worksheet.Cells[3, 4].Value.ToString() : null;
                            company.companyOutPutPath = outputFolder + company.customerName + "\\";
                            if (!Directory.Exists(company.companyOutPutPath))
                            {
                                Directory.CreateDirectory(company.companyOutPutPath);
                            }
                            company.address = worksheet.Cells[1, 4].Value != null ? worksheet.Cells[1, 4].Value.ToString() : null;
                            company.contactPeople = worksheet.Cells[2, 4].Value != null ? worksheet.Cells[2, 4].Value.ToString() : null;

                            company.month = worksheet.Cells[1, 6].Value != null ? DateTime.Parse(worksheet.Cells[1, 6].Value.ToString()).ToString("MM-yyyy") : null;
                            company.invoiceMonth = worksheet.Cells[1, 6].Value != null ? DateTime.Parse(worksheet.Cells[1, 6].Value.ToString()).ToString("MM") : null;



                            if (company.month != null)
                            {
                                var monthSplitList = company.month.Split('-');
                                var lastDay = DateTime.DaysInMonth(Int32.Parse(monthSplitList[1]), Int32.Parse(monthSplitList[0]));
                                company.invoiceDate += lastDay + "-" + company.month;
                                invoiceDate = lastDay + "-" + monthSplitList[0] + "-" + monthSplitList[1];
                                //------------------------Receipt Date
                                DateTime receiptDate = new DateTime(Int32.Parse(monthSplitList[1]), Int32.Parse(monthSplitList[0]), 1).AddMonths(2);
                                var receiptDatelastDay = DateTime.DaysInMonth(receiptDate.Year, receiptDate.Month);
                                company.receiptDate = receiptDatelastDay + "-" + receiptDate.Month + "-" + receiptDate.Year;
                            }
                            company.invoiceNum = worksheet.Cells[2, 6].Value != null ? worksheet.Cells[2, 6].Value.ToString() : null;
                            company.invoiceNoForCompanySalary = worksheet.Cells[3, 6].Value != null ? worksheet.Cells[3, 6].Value.ToString() : null;
                            //date
                            for (int j = 0; j <= cols; j++)
                            {

                                var workDateExecl = worksheet.Cells[6, j].Value != null ? worksheet.Cells[6, j].Value : "";
                                if (workDateExecl.Equals("工作日期"))
                                {
                                    for (int i = 7; i <= rows; i++)
                                    {
                                        if (worksheet.Cells[i, j].Value != null)
                                        {
                                            string tempDate = DateTime.Parse(worksheet.Cells[i, j].Value.ToString()).ToString("dd-MM-yyyy");


                                            worksheet.Cells[i, j].Value = tempDate;
                                            dateList.Add(tempDate);
                                        }
                                    }

                                }
                            }
                            dateList = dateList.Distinct().ToList();

                            //time
                            for (int j = 0; j <= cols; j++)
                            {
                                if (worksheet.Cells[6, j].Value != null)
                                {
                                    if (worksheet.Cells[6, j].Value.Equals("RUSH FEE"))
                                    {
                                        break;
                                    }

                                    switch (worksheet.Cells[6, j].Value)
                                    {
                                        case "站":
                                        case "工作日期":
                                        case "職位":
                                        case "樓層":
                                        case "RUSH FEE":
                                        case "CANCELLATION FEE":
                                        case "REMARKS":
                                        case "TITLE":
                                        case "NAME LIST":
                                        case "更期(特別事項)":
                                        case "ALL REMARKS":
                                        case "日期(特別事項)":
                                        case "Name(特別事項)":
                                        case "":
                                        case "Salary(特別事項)":
                                        case "Reason(特別事項)":
                                            break;
                                        default:
                                            timeList.Add(worksheet.Cells[6, j].Value.ToString());
                                            break;
                                    }

                                }


                            }
                            for (int j = 0; j <= cols; j++)
                            {

                                var oo = worksheet.Cells[6, j].Value;
                                if (oo != null)
                                {
                                    if (oo.Equals("日期(特別事項)"))
                                    {
                                        for (int i = 7; i <= rows; i++)
                                        {
                                            
                                            System.Diagnostics.Debug.WriteLine(i);
                                            var date = worksheet.Cells[i, j].Value != null ? worksheet.Cells[i, j].Value.ToString().Trim() : null;
                                            if (string.IsNullOrEmpty(date))
                                            {
                                                date = null;
                                            }
                                            date = date!=null ? DateTime.Parse(worksheet.Cells[i, j].Value.ToString().Trim()).ToString("dd-MM-yyyy") : null;
                                            var name = worksheet.Cells[i, j + 1].Value != null ? worksheet.Cells[i, j + 1].Value.ToString() : null;
                                            var shift = worksheet.Cells[i, j + 2].Value != null ? worksheet.Cells[i, j + 2].Value.ToString() : null;
                                            var salary = worksheet.Cells[i, j + 3].Value != null ? worksheet.Cells[i, j + 3].Value.ToString() : null;
                                            var reason = worksheet.Cells[i, j + 4].Value != null ? worksheet.Cells[i, j + 4].Value.ToString() : null;

                                          
                                            string[] reasonList = new string[2];
                                            if (reason != null)
                                            {
                                                reasonList = reason.Split(',');
                                                if(reasonList.Length < 2)
                                                {
                                                    throw new Exception("Reason Format 錯, 正常Format : Salary入小時 | Reason 入 OT,原因 / T8,原因\n Salary入金額 Reason 入 bonus,原因");
                                                    
                                                }
                                            }
                                            if (name != null)
                                            {
                                                specialEventsList.Add(new specialEvent
                                                {
                                                    date = date,
                                                    name = name,
                                                    shift = shift,
                                                    hours = salary,
                                                    eventT8orOT = reasonList[0],
                                                    reason = reasonList[1]
                                                });
                                            }

                                        }
                                    }
                                }
                            }

                            for (int j = 0; j <= cols; j++)
                            {
                                for (int e = 0; e < timeList.Count; e++)
                                {
                                    if (timeList[e].Equals(worksheet.Cells[6, j].Value))
                                    {
                                        for (int i = 7; i <= rows; i++)
                                        {
                                            // var dutytime = worksheet.Cells[0, j];
                                            /*    Console.Write(worksheet.Cells[i+1, 1].Value);

                                                Console.Write(worksheet.Cells[i+1, j].Value + " | ");*/

                                            // Console.WriteLine(worksheet.Cells[i + 1, j].Value);
                                            var staffName = worksheet.Cells[i, j].Value != null ? worksheet.Cells[i, j].Value.ToString().Trim() : "";


                                            if (!string.IsNullOrEmpty(staffName))
                                            {
                                                var checkMoreThanOneStaff = staffName.Split(",");
                                                if (checkMoreThanOneStaff.Length > 1)
                                                {
                                                    for (int n = 0; n < checkMoreThanOneStaff.Length; n++)
                                                    {

                                                        tempStaffNameList.Add(checkMoreThanOneStaff[n]);

                                                    }
                                                }
                                                else
                                                {
                                                    tempStaffNameList.Add(staffName);
                                                }

                                            }

                                        }

                                    }
                                }

                            }

                            tempStaffNameList = tempStaffNameList.Distinct().ToList();
                            timeList = timeList.Distinct().ToList();
                            for (int i = 0; i < tempStaffNameList.Count; i++)
                            {
                                staffNameList.Add(new staffList { name = tempStaffNameList[i] });
                            }
                            var titleColumn = 0;
                            for (int j = 0; j <= cols; j++)
                            {
                                if (worksheet.Cells[6, j].Value != null)
                                {
                                    if (worksheet.Cells[6, j].Value.Equals("職位"))
                                    {
                                        titleColumn = j;
                                    }
                                }

                                for (int e = 0; e < timeList.Count; e++)
                                {
                                    if (timeList[e].Equals(worksheet.Cells[6, j].Value))
                                    {
                                        for (int i = 7; i <= rows; i++)
                                        {

                                            for (int g = 0; g < staffNameList.Count; g++)
                                            {

                                                if (worksheet.Cells[i, j].Value != null)
                                                {

                                                    var ExeclRowStaffName = worksheet.Cells[i, j].Value.ToString().Trim();

                                                    var checkMoreThanOneStaff1 = ExeclRowStaffName.Split(",");


                                                    for (int q = 0; q < checkMoreThanOneStaff1.Count(); q++)
                                                    {
                                                        if (staffNameList[g].name.Equals(checkMoreThanOneStaff1[q]))
                                                        {
                                                            //var dutyHours = 
                                                            var date1 = worksheet.Cells[i, 1].Value;
                                                            var title = worksheet.Cells[i, titleColumn].Value;
                                                            var dutyTime = worksheet.Cells[6, j].Value.ToString();

                                                            var hoursAndShiftList = dutyTime.Split('[', ']')[1].Split(",");


                                                            string shift = string.Empty;
                                                            if (hoursAndShiftList.Count() > 1)
                                                            {
                                                                shift = hoursAndShiftList[1];
                                                            }

                                                            staffNameList[g].duty.Add(
                                                                 new duty()
                                                                 {
                                                                     date = date1.ToString(),
                                                                     dutyTime = dutyTime,
                                                                     dutyHours = hoursAndShiftList[0],
                                                                     shift = shift,
                                                                     salary = "",
                                                                     title = title.ToString()
                                                                 }
                                                                 );




                                                        }
                                                    }


                                                }
                                            }
                                        }

                                    }
                                }
                            }




                            //------------------- Read salary execl and insert Salary to every staff -------------------------//
                            Workbook SalaryWb = new Workbook(generalFolder + "staff_salary.xlsx");

                            Worksheet SalarySheet = SalaryWb.Worksheets[0];

                            for (int i = 0; i < SalaryWb.Worksheets.Count; i++)
                            {

                                if (SalaryWb.Worksheets[i].Name == company.customerName + company.invoiceNoForCompanySalary)
                                {
                                    SalarySheet = SalaryWb.Worksheets[i];
                                }
                            }
                            Console.WriteLine("Staff Salary Table : " + SalarySheet.Name);

                            int SalaryRows = SalarySheet.Cells.MaxDataRow;
                            int SalaryCols = SalarySheet.Cells.MaxDataColumn;
                            List<hourSalary> salaryList = new List<hourSalary>();
                            for (int i = 1; i <= SalaryCols; i++)
                            {

                                for (int e = 0; e < SalaryRows; e++)
                                {
                                    var title = SalarySheet.Cells[0, i].Value.ToString();
                                    var hours = SalarySheet.Cells[e + 1, 0].Value.ToString();
                                    var salary = SalarySheet.Cells[e + 1, i].Value == null ? "0" : SalarySheet.Cells[e + 1, i].Value.ToString();
                                    salaryList.Add(new hourSalary
                                    {
                                        title = title,
                                        hours = hours,
                                        salary = salary

                                    });


                                }
                            }

                            Workbook SalaryWb1 = new Workbook(generalFolder + "company_salary.xlsx");

                            Worksheet SalarySheet1 = SalaryWb1.Worksheets[0];

                            for (int i = 0; i < SalaryWb1.Worksheets.Count; i++)
                            {

                                if (SalaryWb1.Worksheets[i].Name == company.customerName + company.invoiceNoForCompanySalary)
                                {
                                    SalarySheet1 = SalaryWb1.Worksheets[i];
                                }
                            }
                            Console.WriteLine("Company Salary Table : " + SalarySheet1.Name);

                            int SalaryRows1 = SalarySheet1.Cells.MaxDataRow;
                            int SalaryCols1 = SalarySheet1.Cells.MaxDataColumn;
                            List<hourSalary> companySalaryList = new List<hourSalary>();

                            for (int i = 1; i <= SalaryCols1; i++)
                            {

                                for (int e = 0; e < SalaryRows1; e++)
                                {
                                    var title = SalarySheet1.Cells[0, i].Value.ToString();
                                    var hours = SalarySheet1.Cells[e + 1, 0].Value.ToString();
                                    var salary = SalarySheet1.Cells[e + 1, i].Value == null ? "0" : SalarySheet1.Cells[e + 1, i].Value.ToString();
                                    companySalaryList.Add(new hourSalary
                                    {
                                        title = title,
                                        hours = hours,
                                        salary = salary

                                    });


                                }
                            }

                            for (int i = 0; i < staffNameList.Count; i++)
                            {

                                for (int q = 0; q < staffNameList[i].duty.Count; q++)
                                {


                                    for (int e = 0; e < companySalaryList.Count; e++)
                                    {
                                        if (companySalaryList[e].hours == staffNameList[i].duty[q].dutyHours && companySalaryList[e].title == staffNameList[i].duty[q].title)
                                        {

                                            staffNameList[i].duty[q].companyUnitPrice = companySalaryList[e].salary;
                                            staffNameList[i].duty[q].companySalary = companySalaryList[e].salary;

                                        }

                                    }
                                }
                            }
                            for (int i = 0; i < staffNameList.Count; i++)
                            {

                                for (int q = 0; q < staffNameList[i].duty.Count; q++)
                                {


                                    for (int e = 0; e < salaryList.Count; e++)
                                    {
                                        if (salaryList[e].hours == staffNameList[i].duty[q].dutyHours && salaryList[e].title == staffNameList[i].duty[q].title)
                                        {

                                            staffNameList[i].duty[q].staffUnitPrice = salaryList[e].salary;
                                            staffNameList[i].duty[q].salary = salaryList[e].salary;

                                        }

                                    }
                                }
                            }
                            for (int i = 0; i < staffNameList.Count; i++)
                            {

                                var titleDutyDistinctList = staffNameList[i].duty.Select(x => new { x.title, x.dutyHours }).Distinct().ToList();

                                for (int e = 0; e < titleDutyDistinctList.Count; e++)
                                {
                                    tittleDuty tittleDuty = new tittleDuty
                                    {
                                        title = titleDutyDistinctList[e].title,
                                        hours = titleDutyDistinctList[e].dutyHours,
                                        dutyList = new List<duty>()
                                    };
                                    staffNameList[i].titleDuty.Add(tittleDuty);

                                }

                                for (int q = 0; q < staffNameList[i].duty.Count; q++)
                                {
                                    for (int z = 0; z < staffNameList[i].titleDuty.Count; z++)
                                    {
                                        if (staffNameList[i].duty[q].dutyHours == staffNameList[i].titleDuty[z].hours && staffNameList[i].duty[q].title == staffNameList[i].titleDuty[z].title)
                                        {
                                            staffNameList[i].titleDuty[z].companyUnitPrice = staffNameList[i].duty[q].companySalary;
                                            staffNameList[i].titleDuty[z].staffUnitPrice = staffNameList[i].duty[q].salary;
                                            staffNameList[i].titleDuty[z].dutyList.Add(staffNameList[i].duty[q]);


                                        }
                                    }
                                }

                            }



                            for (int i = 0; i < staffNameList.Count; i++)
                            {

                                staffNameList[i].duty = staffNameList[i].duty.OrderBy(x => x.date).ToList();

                            }

                            //------------------- Read Bank Record execl  -------------------------//
                            Workbook BankWb = new Workbook(generalFolder + "bank.xlsx");


                            Worksheet BankSheet = BankWb.Worksheets[0];
                            int BankRows = BankSheet.Cells.MaxDataRow;
                            int BankCols = BankSheet.Cells.MaxDataColumn;

                            for (int i = 0; i < BankCols; i++)
                            {
                                for (int e = 1; e < BankRows; e++)
                                {
                                    if (BankSheet.Cells[0, i].Value.ToString() == "中文名")
                                    {
                                        if (BankSheet.Cells[e, i].Value != null)
                                        {
                                         /*   System.Diagnostics.Debug.WriteLine(BankSheet.Cells[e, i].Value.ToString());*/
                                            for (int q = 0; q < staffNameList.Count; q++)
                                            {

                                                if (staffNameList[q].name == BankSheet.Cells[e, i].Value.ToString())
                                                {
                                                    staffNameList[q].bankAccount = BankSheet.Cells[e, 2].Value != null ? BankSheet.Cells[e, 2].Value.ToString() : null;
                                                    staffNameList[q].engName = BankSheet.Cells[e, 1].Value != null ? BankSheet.Cells[e, 1].Value.ToString() : null;
                                                    staffNameList[q].firstRegisterFees = BankSheet.Cells[e, 3].Value != null ? BankSheet.Cells[e, 3].Value.ToString() : null;
                                                    staffNameList[q].uniformFees = BankSheet.Cells[e, 4].Value != null ? BankSheet.Cells[e, 4].Value.ToString() : null;
                                                    staffNameList[q].cancelFees = BankSheet.Cells[e, 5].Value != null ? BankSheet.Cells[e, 5].Value.ToString() : null;
                                                    staffNameList[q].otherFees = BankSheet.Cells[e, 6].Value != null ? BankSheet.Cells[e, 6].Value.ToString() : null;
                                                    staffNameList[q].urgentFees = BankSheet.Cells[e, 7].Value != null ? BankSheet.Cells[e, 7].Value.ToString() : null;
                                                    staffNameList[q].bonus = BankSheet.Cells[e, 8].Value != null ? BankSheet.Cells[e, 8].Value.ToString() : null;
                                                    staffNameList[q].transportFees = BankSheet.Cells[e, 9].Value != null ? BankSheet.Cells[e, 9].Value.ToString() : null;
                                                    staffNameList[q].remark = BankSheet.Cells[e, 11].Value != null ? BankSheet.Cells[e, 11].Value.ToString() : null;

                                                }

                                            }
                                        }

                                    }

                                }
                            }




                            for (int i = 0; i < staffNameList.Count; i++)
                            {
                                /* if (i == 2)
                                 {
                                     var ggw = "";
                                 }*/

                                for (int e = 0; e < staffNameList[i].duty.Count; e++)
                                {

                                    double salary = 0;
                                    double companySalary = 0;
                                    try
                                    {
                                        salary = Convert.ToDouble(staffNameList[i].duty[e].salary);

                                        if (salary == null || salary == 0)
                                        {
                                            throw new Exception(staffNameList[i].duty[e].title + "，" + staffNameList[i].duty[e].dutyTime + "　Staff Salary Not Found");
                                        }
                                    }
                                    catch (Exception)
                                    {

                                        throw new Exception(staffNameList[i].duty[e].title + "，" + staffNameList[i].duty[e].dutyTime + "　Staff Salary Not Found");
                                    }


                                    try
                                    {


                                        companySalary = Convert.ToDouble(staffNameList[i].duty[e].companySalary);

                                        if (companySalary == null || companySalary == 0)
                                        {
                                            throw new Exception(staffNameList[i].duty[e].title + "，" + staffNameList[i].duty[e].dutyTime + "　Company Salary Not Found");
                                        }

                                    }
                                    catch (Exception)
                                    {

                                        throw new Exception(staffNameList[i].duty[e].title + "，" + staffNameList[i].duty[e].dutyTime + "　Company Salary Not Found");
                                    }

                                    decimal t8Staffsalary = 0;
                                    decimal t8CompanySalary = 0;
                                    for (int p = 0; p < specialEventsList.Count; p++)
                                    {
                                        if (specialEventsList[p].name == staffNameList[i].name && specialEventsList[p].date == staffNameList[i].duty[e].date && specialEventsList[p].shift == staffNameList[i].duty[e].dutyTime)
                                        {
                                             
                              
                                            if (specialEventsList[p].eventT8orOT.ToUpper() == "T8")
                                            {
                                                staffNameList[i].duty[e].T8reason = specialEventsList[p].reason;
                                                /*  salary *= 1.5;
                                                  companySalary *= 2;*/
                                                staffNameList[i].duty[e].T8 = true;

                                                staffNameList[i].duty[e].T8StaffRemovesalary = Math.Floor(Convert.ToDouble((salary / (Convert.ToDouble(staffNameList[i].duty[e].dutyHours) * 60)) * Convert.ToDouble(specialEventsList[p].hours)));
                                                staffNameList[i].duty[e].T8StaffAddsalary = Math.Floor(Convert.ToDouble((salary / (Convert.ToDouble(staffNameList[i].duty[e].dutyHours) *60)) * Convert.ToDouble(specialEventsList[p].hours) * 1.5));
                                                staffNameList[i].duty[e].T8CompanyRemoveSalary = Math.Floor(Convert.ToDouble((companySalary / (Convert.ToDouble(staffNameList[i].duty[e].dutyHours) * 60)) * Convert.ToDouble(specialEventsList[p].hours)));
                                                staffNameList[i].duty[e].T8CompanyAddSalary = Math.Floor(Convert.ToDouble((companySalary / (Convert.ToDouble(staffNameList[i].duty[e].dutyHours) * 60)) * Convert.ToDouble(specialEventsList[p].hours) * 2));


                                                staffNameList[i].duty[e].T8CompanySalaryFormula = @$"(正常收費${companySalary} + ${staffNameList[i].duty[e].T8CompanyRemoveSalary}({specialEventsList[p].hours}分鐘))";
                                                staffNameList[i].duty[e].T8StaffSalaryFormula = @$"(正常收費${salary} + ${staffNameList[i].duty[e].T8StaffRemovesalary}({specialEventsList[p].hours}分鐘))";

                                                t8Staffsalary += Convert.ToDecimal(salary) + Convert.ToDecimal(staffNameList[i].duty[e].T8StaffRemovesalary);
                                                t8CompanySalary += Convert.ToDecimal(companySalary) + Convert.ToDecimal(staffNameList[i].duty[e].T8CompanyRemoveSalary);


                                                staffNameList[i].duty[e].t8CompanySalary = Convert.ToDecimal(t8CompanySalary);
                                                staffNameList[i].duty[e].t8StaffSalary = Convert.ToDecimal(t8Staffsalary);

                                                staffNameList[i].totalSalaryForCompany += t8CompanySalary;
                                                staffNameList[i].totalStaffSalary += t8Staffsalary;
                                  /*              staffNameList[i].totalSalaryForCompany +=  Convert.ToDecimal(staffNameList[i].duty[e].T8CompanyAddSalary);
                                                staffNameList[i].totalStaffSalary -= Convert.ToDecimal(staffNameList[i].duty[e].T8StaffRemovesalary);
                                                staffNameList[i].totalSalaryForCompany -= Convert.ToDecimal(staffNameList[i].duty[e].T8CompanyRemoveSalary);
                                                staffNameList[i].totalStaffSalary += Convert.ToDecimal(staffNameList[i].duty[e].T8StaffAddsalary);*/

                                            }
                                            else if (specialEventsList[p].eventT8orOT.ToUpper() == "OT")
                                            {
                                                staffNameList[i].duty[e].OTreason = specialEventsList[p].reason;
                                                staffNameList[i].duty[e].OT = true;

                                                decimal minutes = Convert.ToDecimal(staffNameList[i].duty[e].dutyHours) * 60;
                                                 
                                                decimal revisedSalary = 0;
                                                decimal revisedCompanySalary = 0;
                                                string addOrNo = "+";
                                                if (specialEventsList[p].hours.Contains("-"))
                                                {
                                                    addOrNo = "";
                                                }
                                                if (staffNameList[i].duty[e].T8 == true)
                                                {
                                                    revisedSalary = Convert.ToInt32((Convert.ToDouble(minutes) + Convert.ToDouble(specialEventsList[p].hours)) / Convert.ToDouble(minutes) * Convert.ToInt32(t8Staffsalary));
                                                    revisedCompanySalary = Math.Floor((minutes + Convert.ToInt32(specialEventsList[p].hours)) / minutes * Convert.ToInt32(t8CompanySalary));
                                                    staffNameList[i].duty[e].OTCompanySalaryFormula = @$"({Math.Floor(minutes)}分鐘 {addOrNo} {specialEventsList[p].hours}分鐘) / {Math.Floor(minutes)}分鐘 * T8收費{t8CompanySalary}";
                                                    staffNameList[i].duty[e].OTStaffSalaryFormula = @$"({Math.Floor(minutes)}分鐘 {addOrNo} {specialEventsList[p].hours}分鐘) / {Math.Floor(minutes)}分鐘 * T8收費{t8Staffsalary}";
                                                    staffNameList[i].duty[e].t8CompanySalary = 0;
                                                    staffNameList[i].duty[e].t8StaffSalary = 0;

                                                }
                                                else
                                                {
                                                    revisedSalary = Convert.ToInt32((Convert.ToDouble(minutes) + Convert.ToDouble(specialEventsList[p].hours)) / Convert.ToDouble(minutes) * salary);
                                                    revisedCompanySalary = Math.Floor((minutes + Convert.ToInt32(specialEventsList[p].hours)) / minutes * Convert.ToInt32(companySalary));
                                                    staffNameList[i].duty[e].OTCompanySalaryFormula = @$"({Math.Floor(minutes)}分鐘 {addOrNo} {specialEventsList[p].hours}分鐘) / {Math.Floor(minutes)}分鐘 * 正常收費{companySalary}";
                                                    staffNameList[i].duty[e].OTStaffSalaryFormula = @$"({Math.Floor(minutes)}分鐘 {addOrNo} {specialEventsList[p].hours}分鐘) / {Math.Floor(minutes)}分鐘 * 正常收費{salary}";

                                                }




                                                staffNameList[i].duty[e].salary = Convert.ToDecimal(salary).ToString();
                                                staffNameList[i].duty[e].companySalary = Convert.ToDecimal(companySalary.ToString()).ToString();

                                                staffNameList[i].duty[e].OTcompanySalary = Convert.ToDecimal(revisedCompanySalary);
                                                staffNameList[i].duty[e].OTStaffsalary = Convert.ToDecimal(revisedSalary);


                                                staffNameList[i].totalSalaryForCompany += Convert.ToDecimal(revisedCompanySalary);
                                                staffNameList[i].totalStaffSalary += Convert.ToDecimal(revisedSalary);

                                            }
                                            else if(specialEventsList[p].eventT8orOT.ToUpper() == "BONUS")
                                            {
                                                staffNameList[i].duty[e].BonusReason = specialEventsList[p].reason;
                                                staffNameList[i].duty[e].bonus = true;
                                                staffNameList[i].duty[e].bonusSalary = decimal.Parse(specialEventsList[p].hours);
                                                staffNameList[i].totalSalaryForCompany += decimal.Parse(specialEventsList[p].hours);
                                                staffNameList[i].totalStaffSalary += decimal.Parse(specialEventsList[p].hours); 

                                            }
                                        }
                                    }

                                    if(staffNameList[i].duty[e].OT == false && staffNameList[i].duty[e].T8 == false)
                                    {
                                        staffNameList[i].totalSalaryForCompany += Convert.ToDecimal(companySalary);
                                        staffNameList[i].totalStaffSalary += Convert.ToDecimal(salary);
                                    }
                                    else
                                    {
                                     

                                    }
                                

                                }

                                staffNameList[i].totalSalaryOld = staffNameList[i].totalStaffSalary;

                                if (staffNameList[i].firstRegisterFees != null)
                                {

                                    staffNameList[i].totalStaffSalary += Convert.ToDecimal(staffNameList[i].firstRegisterFees);

                                }
                                if (staffNameList[i].uniformFees != null)
                                {


                                    staffNameList[i].totalStaffSalary += Convert.ToDecimal(staffNameList[i].uniformFees);

                                }
                                if (staffNameList[i].cancelFees != null)
                                {


                                    staffNameList[i].totalStaffSalary += Convert.ToDecimal(staffNameList[i].cancelFees);

                                }
                                if (staffNameList[i].urgentFees != null)
                                {


                                    staffNameList[i].totalStaffSalary += Convert.ToDecimal(staffNameList[i].urgentFees);

                                }
                                if (staffNameList[i].bonus != null)
                                {


                                    staffNameList[i].totalStaffSalary += Convert.ToDecimal(staffNameList[i].bonus);

                                }
                                if (staffNameList[i].transportFees != null)
                                {


                                    staffNameList[i].totalStaffSalary += Convert.ToDecimal(staffNameList[i].transportFees);

                                }
                            }


                            company.staffLists = staffNameList;
                            for (int i = 0; i < staffNameList.Count; i++)
                            {
                                bool exist = false;
                                companyDuty companyDuty = new companyDuty
                                {
                                    companyName = company.customerName,
                                    duty = staffNameList[i].duty,
                                    titleDuty = staffNameList[i].titleDuty
                                };



                                for (int q = 0; q < allStaffList.Count; q++)
                                {
                                    if (allStaffList[q].name == staffNameList[i].name)
                                    {
                                        exist = true;
                                        allStaffList[q].companyDuty.Add(companyDuty);
                                    }

                                }
                                if (exist == false)
                                {
                                    allStaff staff = new allStaff
                                    {
                                        name = staffNameList[i].name,

                                        engName = staffNameList[i].engName,
                                        firstRegisterFees = staffNameList[i].firstRegisterFees,
                                        cancelFees = staffNameList[i].cancelFees,
                                        bankAccount = staffNameList[i].bankAccount,

                                        otherFees = staffNameList[i].otherFees,
                                        remark = staffNameList[i].remark,
                                        uniformFees = staffNameList[i].uniformFees,
                                        urgentFees = staffNameList[i].urgentFees,
                                        bonus = staffNameList[i].bonus,
                                        transportFees = staffNameList[i].transportFees
                                    };
                                    staff.companyDuty.Add(companyDuty);
                                    allStaffList.Add(staff);

                                }
                            }

                            companyList.Add(company);
                        }



                        await company_invoice(companyList);
                        // Console.WriteLine("company invoice processing");
                        for (int o = 0; o < allStaffList.Count; o++)
                        {
                            for (int p = 0; p < allStaffList[o].companyDuty.Count; p++)
                            {
                                for (int q = 0; q < companyList.Count; q++)
                                {
                                    for (int i = 0; i < companyList[q].staffLists.Count; i++)
                                    {
                                        if (allStaffList[o].companyDuty[p].companyName == companyList[q].customerName)
                                        {
                                            if (allStaffList[o].name == companyList[q].staffLists[i].name)
                                            {
                                                allStaffList[o].companyDuty[p].title = companyList[q].staffLists[i].title;
                                                allStaffList[o].companyDuty[p].pdfDescription = companyList[q].staffLists[i].pdfDescription;
                                                allStaffList[o].companyDuty[p].oldTotalSalary = companyList[q].staffLists[i].totalSalaryOld;
                                            }

                                        }
                                    }
                                }

                            }
                        }
                        // Console.WriteLine("company invoice processing Done");
                        /*                 Console.WriteLine("要staff invoice 輸入 1, 不要staff invoice 輸入 2");
                                         string check2 = Console.ReadLine();*/

                        await allStaffInvoice(allStaffList);
                        Console.WriteLine("BankAccount Execl Processing");
                        bankAccount(allStaffList);
                        Console.WriteLine("BankAccount Execl Done");
                        Console.WriteLine("Company History Execl Processing");
                        companyHistory(companyList);
                        Console.WriteLine("Company History Execl Done");
                        Console.WriteLine("TotalAmount Execl Processing");
                        companyTotalAmountList.totalAmount = companyTotalAmountList.eachTotal.Sum(e => e.total);
                        staffTotalAmountList.totalAmount = staffTotalAmountList.eachTotal.Sum(e => e.total);
                        totalAmount(companyTotalAmountList, staffTotalAmountList);
                        Console.WriteLine("TotalAmount Execl Done");

                    }
                }
                else
                {
                    Console.WriteLine("Unable to Access~~~");
                    Console.ReadLine();
                }



            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.ReadLine();

            }
        }

        public static string totalAmount(totalAmountObj companyAmountList, totalAmountObj staffAmountList)
        {
            Workbook wb = new Workbook();

            // 得到第一個工作表。
            Worksheet sheet1 = wb.Worksheets[0];
            sheet1.Cells.SetColumnWidth(1, 20.0);
            sheet1.Cells.SetColumnWidth(2, 30.0);
            sheet1.Cells.SetColumnWidth(3, 10.0);
            sheet1.Cells.SetColumnWidth(4, 10.0);
            sheet1.Cells.SetColumnWidth(5, 10.0);
            sheet1.Cells.SetColumnWidth(6, 10.0);
            sheet1.Cells.SetColumnWidth(7, 10.0);
            // 獲取工作表的單元格集合
            Cells cells = sheet1.Cells;

            // 為單元格設置值
            Aspose.Cells.Cell cell = cells["A2"];
            cell.PutValue("院舍");
            cell = cells["B2"];
            cell.PutValue("Total");


            cell = cells["E2"];
            cell.PutValue("員工");
            cell = cells["F2"];
            cell.PutValue("Total");

            cell = cells[@$"B1"];
            cell.PutValue(companyAmountList.totalAmount);

            cell = cells[@$"F1"];
            cell.PutValue(staffTotalAmountList.totalAmount);

            for (int i = 0; i < companyAmountList.eachTotal.Count; i++)
            {
                cell = cells[@$"A{i + 3}"];
                cell.PutValue(companyAmountList.eachTotal[i].name);
                cell = cells[@$"B{i + 3}"];
                cell.PutValue(companyAmountList.eachTotal[i].total);
            }

            for (int i = 0; i < staffAmountList.eachTotal.Count; i++)
            {
                cell = cells[@$"E{i + 3}"];
                cell.PutValue(staffAmountList.eachTotal[i].name);
                cell = cells[@$"F{i + 3}"];
                cell.PutValue(staffAmountList.eachTotal[i].total);
            }
            // 保存 Excel 文件。
            wb.Save(outputFolder + "Total Amount_output.xlsx", SaveFormat.Xlsx);

            return "";
        }
        public static string bankAccount(List<allStaff> allstaffList)
        {
            Workbook wb = new Workbook();

            // 得到第一個工作表。
            Worksheet sheet1 = wb.Worksheets[0];
            sheet1.Cells.SetColumnWidth(1, 20.0);
            sheet1.Cells.SetColumnWidth(2, 30.0);
            sheet1.Cells.SetColumnWidth(3, 10.0);
            sheet1.Cells.SetColumnWidth(4, 10.0);
            sheet1.Cells.SetColumnWidth(5, 10.0);
            sheet1.Cells.SetColumnWidth(6, 10.0);
            sheet1.Cells.SetColumnWidth(7, 10.0);
            // 獲取工作表的單元格集合
            Cells cells = sheet1.Cells;

            // 為單元格設置值
            Aspose.Cells.Cell cell = cells["A1"];
            cell.PutValue("中文名");
            cell = cells["B1"];
            cell.PutValue("ENG NAME");
            cell = cells["C1"];
            cell.PutValue("銀行");
            cell = cells["D1"];
            cell.PutValue("首次登記費");
            cell = cells["E1"];
            cell.PutValue("制服費");
            cell = cells["F1"];
            cell.PutValue("取消費");
            cell = cells["G1"];
            cell.PutValue("雜費");
            cell = cells["H1"];
            cell.PutValue("加急費");
            cell = cells["I1"];
            cell.PutValue("獎金");
            cell = cells["J1"];
            cell.PutValue("交通費");
            cell = cells["K1"];
            cell.PutValue("TOTAL");
            cell = cells["L1"];
            cell.PutValue("REMARK");

            for (int i = 0; i < allstaffList.Count; i++)
            {
                cell = cells[@$"A{i + 2}"];
                cell.PutValue(allstaffList[i].name);
                cell = cells[@$"B{i + 2}"];
                cell.PutValue(allstaffList[i].engName);
                cell = cells[@$"C{i + 2}"];

                cell.PutValue(allstaffList[i].bankAccount);
                cell = cells[@$"D{i + 2}"];
                cell.PutValue(allstaffList[i].firstRegisterFees);
                cell = cells[@$"E{i + 2}"];
                cell.PutValue(allstaffList[i].uniformFees);
                cell = cells[@$"F{i + 2}"];
                cell.PutValue(allstaffList[i].cancelFees);
                cell = cells[@$"G{i + 2}"];
                cell.PutValue(allstaffList[i].otherFees);
                cell = cells[@$"H{i + 2}"];
                cell.PutValue(allstaffList[i].urgentFees);
                cell = cells[@$"I{i + 2}"];
                cell.PutValue(allstaffList[i].bonus);
                cell = cells[@$"J{i + 2}"];
                cell.PutValue(allstaffList[i].transportFees);
                cell = cells[@$"K{i + 2}"];
                cell.PutValue(allstaffList[i].totalSalary);
                cell = cells[@$"L{i + 2}"];
                cell.PutValue(allstaffList[i].remark);
            }

            // 保存 Excel 文件。
            wb.Save(outputFolder + "System Bank Record_Program_output.xlsx", SaveFormat.Xlsx);

            return "";
        }
        public static async Task<string> company_invoice(List<company> companyList)
        {
            var renderer = new HtmlToPdf();
            if (check1 == "1")
            {
                Console.WriteLine("company invoice processing");
            }
            for (int q = 0; q < companyList.Count; q++)
            {
                if (check1 == "1")
                {
                    Console.WriteLine("company invoice " + q + "/" + companyList.Count);
                }
                decimal allTotal = 0;
                renderer.PrintOptions.Title = companyList[q].companyOutPutPath + companyList[q].customerName + "總invoice_" + companyList[q].invoiceMonth + "月";
                string body = string.Empty;

                for (int i = 0; i < companyList[q].staffLists.Count; i++)
                {

                    allTotal += companyList[q].staffLists[i].totalSalaryForCompany;


                    for (int e = 0; e < companyList[q].staffLists[i].titleDuty.Count; e++)
                    {
                        //duty order by date asc
                        companyList[q].staffLists[i].titleDuty[e].dutyList = companyList[q].staffLists[i].titleDuty[e].dutyList.OrderBy(x => x.date).ToList();

                        decimal eachDutytotal = 0;
                        string description = string.Empty;
                        string title = string.Empty;
                        string companyUnitprice = companyList[q].staffLists[i].titleDuty[e].companyUnitPrice;
                        title = companyList[q].staffLists[i].titleDuty[e].title;

                        List<duty> T8List = new List<duty>();
                        List<duty> OTList = new List<duty>();
                        List<duty> BonusList = new List<duty>();
                        T8List = companyList[q].staffLists[i].titleDuty[e].dutyList.Where(x => x.T8 == true).ToList();
                        OTList = companyList[q].staffLists[i].titleDuty[e].dutyList.Where(x => x.OT == true).ToList();
                        BonusList = companyList[q].staffLists[i].titleDuty[e].dutyList.Where(x => x.bonus == true).ToList();
                        int dutyNormalCount = 0;

                        for (int o = 0; o < companyList[q].staffLists[i].titleDuty[e].dutyList.Count; o++)
                        {
                            if (companyList[q].staffLists[i].titleDuty[e].dutyList[o].T8 == false && companyList[q].staffLists[i].titleDuty[e].dutyList[o].OT == false) {
                                dutyNormalCount++;
                                
                                description += DateTime.ParseExact(companyList[q].staffLists[i].titleDuty[e].dutyList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + companyList[q].staffLists[i].titleDuty[e].dutyList[o].shift + ")";
                               
                                if (o == companyList[q].staffLists[i].titleDuty[e].dutyList.Count - 1)
                                {

                                }
                                else
                                {
                                    description += ", ";

                                }
                           

                                eachDutytotal += decimal.Parse(companyList[q].staffLists[i].titleDuty[e].dutyList[o].companySalary);
                                }
                        }

                        if (!string.IsNullOrEmpty(description))
                        {

                            body += @$"<tr>
                           <td style= 'font-family: verdana'>{companyList[q].staffLists[i].name}</td>
                           <td>{title}</td>
                           <td>{description}</td>
                           <td style='text-align: right;'>{companyList[q].staffLists[i].titleDuty[e].companyUnitPrice}</td>
                           <td style='text-align: center;'>{dutyNormalCount}</td>
                           <td style='text-align: right;'>{eachDutytotal}</td>
                         </tr>";
                        }


                        for (int o = 0; o < BonusList.Count; o++)
                        {
                            body += @$"<tr>
                           <td style= 'font-family: verdana'>{companyList[q].staffLists[i].name}</td>
                           <td>{title}</td>
                           <td>{DateTime.ParseExact(BonusList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + BonusList[o].shift + ")"}{BonusList[o].BonusReason}</td>
                           <td style='text-align: right;'>{BonusList[o].bonusSalary}</td>
                           <td style='text-align: center;'>1</td>
                           <td style='text-align: right;'>{BonusList[o].bonusSalary}</td>
                         </tr>";

                            eachDutytotal += BonusList[o].bonusSalary;
                        }
                        for (int o = 0; o < T8List.Count; o++)
                        {
                            body += @$"<tr>
                           <td style= 'font-family: verdana'>{companyList[q].staffLists[i].name}</td>
                           <td>{title}</td>
                           <td>{DateTime.ParseExact(T8List[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + T8List[o].shift + ")"}{T8List[o].T8reason}{T8List[o].T8CompanySalaryFormula}</td>
                           <td style='text-align: right;'>{T8List[o].t8CompanySalary}</td>
                           <td style='text-align: center;'>1</td>
                           <td style='text-align: right;'>{T8List[o].t8CompanySalary}</td>
                         </tr>";

                            eachDutytotal += T8List[o].t8CompanySalary;
                        }

                        for (int o = 0; o < OTList.Count; o++)
                        {
                            body += @$"<tr>
                           <td style= 'font-family: verdana'>{companyList[q].staffLists[i].name}</td>
                           <td>{title}</td>
                           <td>{DateTime.ParseExact(OTList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + OTList[o].shift + ")"}{OTList[o].OTreason}{OTList[o].OTCompanySalaryFormula}</td>
                           <td style='text-align: right;'>{OTList[o].OTcompanySalary}</td>
                           <td style='text-align: center;'>1</td>
                           <td style='text-align: right;'>{OTList[o].OTcompanySalary}</td>
                         </tr>";

                            eachDutytotal += decimal.Parse(OTList[o].OTcompanySalary.ToString());
                        }

                    }
                }

                companyTotalAmountList.eachTotal.Add(new eachTotal { name = companyList[q].customerName + companyList[q].invoiceNoForCompanySalary, total = allTotal });
                var html = $@"<!DOCTYPE html>
        <html>
          {style}
          <body>
            <div class='center'>
              <div>
                <img  src=""{generalFolder}logo.jpg"" width=300 height=100 style='float: left;  margin-bottom: 30px;' />
                <div style='float: right;font-size: 11px;text-align:right'>Room 12, 6/F, Good Harvest Industrial Building, <br> 9 Tsun Wen Road, Tuen Mun <br> 新界屯門震寰路9號好收成工業大廈6樓12室 <br> Tel 電話號碼 : 3618 9330 <br> Fax no. 傳真號碼 : 3020 1710 <br> Email : info@hygienefirstgroup.com </div>
              </div>
              <table height='30px'>
                <tr>
                  <td style='text-align: center; '>
                    <b>Invoice</b>
                  </td>
                </tr>
              </table>
               <table height='30px'>
                <tr>
                  <td style='text-align: left; '>
                    <b>Bill To:</b>
                  </td>
                </tr>
              </table>
              <table>
                <tr>
                  <td width= '421px'><b>Client Name:</b> {companyList[q].customerName} <br><b>Address:</b> {companyList[q].address} <br><b>Tel:</b> {companyList[q].contactPeople} </td>
                  <td><b>Date:</b> <div style='float: right;'>{companyList[q].invoiceDate}</div>
                    <br><b>Invoice No.</b> <div style='float: right;'>{companyList[q].invoiceNum}</div>
                   
                  </td>
                </tr>
              </table>
              <table>
                <tr>
                  <td width='65px'>Name</td>
                  <td>Title</td>
                  <td>Description</td>
                  <td>Unit Price HK$</td>
                  <td>Qty</td>
                  <td>Total Amount HK$</td>
                </tr>
                {body}
                <tr>
                  <td colspan=""5"" style='text-align: right;'>Total HK$</td>
                  <td style='text-align: right;'>{allTotal}</td>
                </tr>
              </table>
              <table>
                <tr>
                  <td style='font-size: 11px;font-family: verdana'><b>Remarks:</b> <br>1.This payment is now due.  Please settle the payment as soon as possible.  Cheque should be payable to  ‘<b>Hygiene First Company Limited</b>’.<br>
                    2.The amount may also be directly deposited into our bank account. <b>Account Number: 012-742-2-019880-4 </b>(Bank of China Hong Kong).<br>
                    3.Please email to <b>ao@hygienefirstgroup.com</b> or Whatsapp to <b>9326 7321</b> the bank slip to us for our checking. For billing enquiries, please contact our Accounting Department at <b>3618 9333 (Ms Ng/Mr Chan)</b>.
                    

                </td>
                </tr>
                <tr>
                  <td style='font-size: 11px;font-family: verdana'><b> Late Payment Surcharge : </b><br> 1.Should bills remain unpaid after 30 days after the postmark date on the envelope, a 5% surcharge will be added to the outstanding amount.<br>
                    2.After <b>45 days</b>, a 10% interest will be imposed on the outstanding amount.
                </td>
                </tr>
              </table>
              <br>
              <b>For and on behalf of <br> Hygiene First Company Limited </b>
            </div>
          </body>
        </html>";

                //----------------------------------------------------------------------------------------------
                var giveBackCompanyHtml = $@"<!DOCTYPE html>
        <html>
          {style}
          <body>
            <div class='center'>
              <div>
                <img  src=""{generalFolder}logo.jpg"" width=300 height=100 style='float: left;  margin-bottom: 30px;' />
                <div style='float: right;font-size: 11px;text-align:right'>Room 12, 6/F, Good Harvest Industrial Building, <br> 9 Tsun Wen Road, Tuen Mun <br> 新界屯門震寰路9號好收成工業大廈6樓12室 <br> Tel 電話號碼 : 3618 9330 <br> Fax no. 傳真號碼 : 3020 1710 <br> Email : info@hygienefirstgroup.com </div>
              </div>
              <table height='30px'>
                <tr>
                  <td style='text-align: center; '>
                    <b>Receipt</b>
                  </td>
                </tr>
              </table> 
             <table>
                <tr>
                      <td width= '421px'><b>Client Name: {companyList[q].customerName} </b><br><b>Address: {companyList[q].address} </b><br><b>Tel: {companyList[q].contactPeople} </b></td>
                  <td><b>Date:</b> <div style='float: right;'>{companyList[q].receiptDate}</div>
                 
                       <br><b>Status:</b> <div style='float: right;'>Paid</div>
                      
                  </td>
                </tr>
              </table>
              <table>
                <tr>  

                  <td width ='85px'><center>Details</center></td>
                  <td width ='50px'><center>Total Amount Received HK$</center></td>
          
                </tr>
                <tr>
                        <td><center>Invoice No. {companyList[q].invoiceNum}</center></td> 
                          <td><center>HK$ {allTotal}</center></td> 
                      
                        </tr> 
                       
                <tr>
                  <td colspan=""1"" style='text-align: right;'></td>
                  <td style='text-align: right;'><br></td>
                </tr>
              </table>
     
            <table>

            <tr>
                <td style='font-size: 13px;'> <u>Remarks：</u> 
                    <br> Above is the <b>offical receipt</b> for the <b>corresponding invoice</b>. For any receipt enquiries, please contact our Accounting Department (<b>Ms Ng</b>/<b>Mr Chan</b>).                                             <br>Email: ao@hygienefirstgroup.com
                        <br>Office Ext: 3618 9333<br>
                        Mobile / Whatsapp: 9326 7321
                  
</td>
            </tr>
        </table>
        <table height='30px'>
                <tr>
                  <td style='text-align: center; '>
                    <b>Thank You For Choosing Our Service</b>
                  </td>
                </tr>
              </table> 
              <b>For and on behalf of <br> Hygiene First Company Limited </b>
            </div>
          </body>
        </html>";
                companyList[q].invoiceTotalSalary = allTotal.ToString();
                if (check1 == "1")
                {
                    var pdf = await renderer.RenderHtmlAsPdfAsync(html);

                    pdf.SaveAs(companyList[q].companyOutPutPath + companyList[q].customerName + companyList[q].invoiceNoForCompanySalary + "總invoice_" + companyList[q].invoiceMonth + "月" + ".pdf");

                    var forCompanyReceiptpdf = await renderer.RenderHtmlAsPdfAsync(giveBackCompanyHtml);

                    forCompanyReceiptpdf.SaveAs(companyList[q].companyOutPutPath + "Receipt_" + companyList[q].customerName + companyList[q].invoiceNoForCompanySalary + "總invoice_" + companyList[q].invoiceMonth + "月" + ".pdf");

                    using (var sw = new StreamWriter(outputCompanyInvoiceHtmlFolder + companyList[q].customerName + companyList[q].invoiceNoForCompanySalary + "總invoice_" + companyList[q].invoiceMonth + "月.txt"))
                    {
                        sw.WriteLine(html);
                    }

                    using (var sw = new StreamWriter(outputCompanyInvoiceReceiptHtmlFolder + "Receipt_" + companyList[q].customerName + companyList[q].invoiceNoForCompanySalary + "總invoice_" + companyList[q].invoiceMonth + "月.txt"))
                    {
                        sw.WriteLine(giveBackCompanyHtml);
                    }
                }

            }


            return "";
        }

        public static async Task<string> allStaffInvoice(List<allStaff> allStaffList)
        {

            for (int i = 0; i < allStaffList.Count; i++)
            {
                decimal totalSalary = 0;
                string body = string.Empty;

                if (check1 == "2")
                {
                    Console.WriteLine(i + 1 + "/" + allStaffList.Count);
                    Console.WriteLine(allStaffList[i].name + " Processing");
                }
                var renderer = new HtmlToPdf();

                for (int q = 0; q < allStaffList[i].companyDuty.Count; q++)
                {



                    renderer.PrintOptions.Title = allStaffList[i].name;
                    
                    for (int e = 0; e < allStaffList[i].companyDuty[q].titleDuty.Count; e++)
                    {
                        int dutyNormalCount = 0;
                        string description = string.Empty;
                        decimal eachDutytotal = 0;
                        List<duty> T8List = new List<duty>();
                        List<duty> OTList = new List<duty>();
                        List<duty> BonusList = new List<duty>();
                        T8List = allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Where(x => x.T8 == true).ToList();
                        OTList = allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Where(x => x.OT == true).ToList();
                        BonusList = allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Where(x => x.bonus == true).ToList();
                        //duty order by date asc
                        allStaffList[i].companyDuty[q].titleDuty[e].dutyList = allStaffList[i].companyDuty[q].titleDuty[e].dutyList.OrderBy(x => x.date).ToList();

                        for (int o = 0; o < allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Count; o++)
                        {
                           
                            if (allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].T8 == false && allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].OT == false)
                            {
                                dutyNormalCount++;
                                description += DateTime.ParseExact(allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].shift + ")";

                               

                                if (o == allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Count - 1)
                                {

                                }
                                else
                                {
                                    description += ", ";

                                }

                                eachDutytotal += decimal.Parse(allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].salary);
                                 
                            }
                        }

                        if (!string.IsNullOrEmpty(description))
                        {
                            body += @$"<tr>
                         <td>{allStaffList[i].companyDuty[q].companyName}</td>
                          <td>{allStaffList[i].name}</td>
                          <td>{allStaffList[i].companyDuty[q].titleDuty[e].title}</td>
                          <td>{description}</td>
                        <td>{allStaffList[i].companyDuty[q].titleDuty[e].staffUnitPrice}</td>
                        <td style = 'text-align: center;'>{dutyNormalCount}</td>
                        <td style = 'text-align: right;'>{eachDutytotal}</td>
                        </tr> 
                       ";
                           
                        }


                        for (int o = 0; o < BonusList.Count; o++)
                        {
                            body += @$"<tr>
                            <td>{allStaffList[i].companyDuty[q].companyName}</td>
                              <td>{allStaffList[i].name}</td>
                             <td>{allStaffList[i].companyDuty[q].titleDuty[e].title}</td>
                           <td>{DateTime.ParseExact(BonusList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + BonusList[o].shift + ")"}{BonusList[o].BonusReason}</td>
                           <td>{BonusList[o].bonusSalary}</td>
                           <td style='text-align: center;'>1</td>
                           <td style='text-align: right;'>{BonusList[o].bonusSalary}</td>
                         </tr>";

                            eachDutytotal += BonusList[o].bonusSalary;
                        }


                        for (int o = 0; o < T8List.Count; o++)
                        {
                            body += @$"<tr>
                            <td>{allStaffList[i].companyDuty[q].companyName}</td>
                              <td>{allStaffList[i].name}</td>
                             <td>{allStaffList[i].companyDuty[q].titleDuty[e].title}</td>
                           <td>{DateTime.ParseExact(T8List[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + T8List[o].shift + ")"}{T8List[o].T8reason}{T8List[o].T8StaffSalaryFormula}</td>
                           <td>{T8List[o].t8StaffSalary}</td>
                           <td style='text-align: center;'>1</td>
                           <td style='text-align: right;'>{T8List[o].t8StaffSalary}</td>
                         </tr>";

                            eachDutytotal += T8List[o].t8StaffSalary;
                        }

                        for (int o = 0; o < OTList.Count; o++)
                        {
                            body += @$"<tr>
                            <td>{allStaffList[i].companyDuty[q].companyName}</td>
                              <td>{allStaffList[i].name}</td>
                             <td>{allStaffList[i].companyDuty[q].titleDuty[e].title}</td>
                           <td>{DateTime.ParseExact(OTList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + OTList[o].shift + ")"}{OTList[o].OTreason}{OTList[o].OTStaffSalaryFormula}</td>
                           <td>{OTList[o].OTStaffsalary}</td>
                           <td style='text-align: center;'>1</td>
                           <td style='text-align: right;'>{OTList[o].OTStaffsalary}</td>
                         </tr>";

                            eachDutytotal += OTList[o].OTStaffsalary;
                        }

                        totalSalary += eachDutytotal;






                    }



                }
                if (allStaffList[i].cancelFees != null)
                {
                    totalSalary += decimal.Parse(allStaffList[i].cancelFees);


                    body += @$"<tr>
                   <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>取消費</td>
                   <td></td>
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].cancelFees}</td>
                 </tr>";
                }
                if (allStaffList[i].firstRegisterFees != null)
                {
                    totalSalary += decimal.Parse(allStaffList[i].firstRegisterFees);

                    body += @$"<tr>
                   <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>首次登記費</td>
                   <td></td>
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].firstRegisterFees}</td>
                 </tr>";
                }
                if (allStaffList[i].uniformFees != null)
                {

                    totalSalary += decimal.Parse(allStaffList[i].uniformFees);
                    body += @$"<tr>
                   <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>制服費</td>
                   <td></td>
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].uniformFees}</td>
                 </tr>";

                }
                if (allStaffList[i].otherFees != null)
                {

                    totalSalary += decimal.Parse(allStaffList[i].otherFees);
                    body += @$"<tr>
                    <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>雜費</td>
                   <td></td>
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].otherFees}</td>
                 </tr>";

                }

                if (allStaffList[i].urgentFees != null)
                {

                    totalSalary += decimal.Parse(allStaffList[i].urgentFees);
                    body += @$"<tr>
                    <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>加急費</td>
                   <td></td>
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].urgentFees}</td>
                 </tr>";

                }

                if (allStaffList[i].bonus != null)
                {

                    totalSalary += decimal.Parse(allStaffList[i].bonus);
                    body += @$"<tr>
                    <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>獎金</td>
                   <td></td>
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].bonus}</td>
                 </tr>";

                }
                if (allStaffList[i].transportFees != null)
                {

                    totalSalary += decimal.Parse(allStaffList[i].transportFees);
                    body += @$"<tr>
                    <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>交通費</td>
                   <td></td>
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].transportFees}</td>
                 </tr>";

                }


                allStaffList[i].totalSalary = totalSalary.ToString();

                staffTotalAmountList.eachTotal.Add(new eachTotal { name = allStaffList[i].name, total = totalSalary });


                if (check1 == "2")
                {
                    string html = $@" <!DOCTYPE html>
        <html>
          {style}
          <body>
            <div class='center'>
              <div>
                <img  src=""{generalFolder}logo.jpg"" width=300 height=100 style='float: left;  margin-bottom: 30px;' />
                <div style='float: right;font-size: 11px;text-align:right'>Room 12, 6/F, Good Harvest Industrial Building, <br> 9 Tsun Wen Road, Tuen Mun <br> 新界屯門震寰路9號好收成工業大廈6樓12室 <br> Tel 電話號碼 : 3618 9330 <br> Fax no. 傳真號碼 : 3020 1710 <br> Email : info@hygienefirstgroup.com </div>
              </div>
              <table height='30px'>
                <tr>
                  <td style='text-align: center; '>
                    <b>自僱人士服務紀錄表</b>
                  </td>
                </tr>
              </table> 
             <table>
                <tr>
                  <td  width ='499px'><b>備注：</b><br><b>1.自取現金支票：</b> 行政費用HKD$20/次。(地址：屯門震寰路九號好收成工業大廈10樓04室）<br><b>2.郵寄現金支票: </b>行政費用HKD$30/次 (平郵/郵局掛號 另加HKD$16/順豐到付另加HKD$18)，郵寄風險自負。<br><b>3.首次登記:</b>自僱人士獲派首次配對服務後，本公司將收取一次性HKD$50為首次登記費用。<br><b>4.更改服務酬金方式：</b> 行政費用HKD$20/次。           </td>
                  <td><div>服務酬金結算日：<br>{invoiceDate}</div>
                
                   
                  </td>
                </tr>
              </table>
              <table>
                <tr>  

                  <td  width ='100px'>服務地點</td>
                  <td width ='50px'>自僱人士名稱</td>
                  <td width ='40px'>稱謂</td>
                  <td  width ='250px'>服務日期</td>
                    <td width ='55px'>服務酬金</td>
                  <td width ='30px'>服務次數</td>
                  <td  width ='100px'>總服務酬金(HKD)</td>
                </tr>
                {body}
                <tr>
                  <td colspan=""6"" style='text-align: right;'>總服務酬金(HKD)</td>
                  <td style='text-align: right;'>{totalSalary}</td>
                </tr>
              </table>
              <table>
                
                <tr>
                  <td style='font-size: 16px;font-family: verdana'><center><I> 如對以上服務紀錄表有任何查詢，請 <b>Whatsapp 6791 6812</b> 與會計部同事聯絡。  </I></center></td>
                </tr>
              </table>
              <br>
            <table>

            <tr>
                <td style='font-size: 13px;font-family: verdana'> <u>政策需知：</u> 
                    <br> 1.本公司與院舍商討更其後有權對服務作出修正、增加、刪除。 
                    <br> 2.如服務時間 ≥7 小時， 院舍會提供最少 30 至 60 分鐘用膳時間（最終由院舍決定）。
                    <br> 3.以下公眾假期的服務費用將按標準費用的 1.5 倍或 2 倍支付: 中秋正日 ( 1.5 倍)，冬至、農曆新年前夕、農曆年初一、二、三 ( 2 倍)。 
                    <br> 4.八號烈風或暴風信號 或 黑色暴雨警告信號 懸掛期間之服務費用為標準費用的 1.5 倍。護理人員如在下班時仍然懸掛八號烈風或暴風信號，可獲最多$100交通津貼，實報實銷。 (以每滿半小時為計算單位)
                    <br> 5. 如接單後出現甩更、遲到、病假等，所有勤工獎金一律取消，及其優先安排工作的次序會延後。
                    <br><br><u>服務酬金需知： </u><br> 1.如有遲到，遲到之鐘數，將由正常工作時間扣除，不獲計算工資。 <br> 2.本公司每月18-22號內轉帳上月服務酬金。 <br><br>  <u>告假須知：</u> <br> 1.如需告假，請提早三個工作天(72小時)通知。
                    <br> 2.不可自行取消由醫護服務選項，如有特殊情況，請馬上通知。如在本公司辦公以外時間(0900-2100)遇上突發事件、或需要臨時請假 請務必致電: 9044 3186 / 9502 4162 / 60862287。<br> 3.即日請假需盡快通知我們，並必須自行致電院舍請假(院舍資料/電話可看訊息上列）。
                    <br> 4.如少於48小時內通知我們，須付港幣200元作行政費用。 
                    <br> 5.如少於12小時內內通知我們，須付港幣300元作行政費用。
                     <br> 6.如有提供服務當日醫生證明書(病假紙)可獲豁免行政費。
</td>
            </tr>
        </table>
              <b>For and on behalf of <br> Hygiene First Company Limited </b>
            </div>
          </body>
        </html>";
                    var pdf = await renderer.RenderHtmlAsPdfAsync(html);

                    pdf.SaveAs(outputFolder + "\\staff_invoice\\" + $@"{allStaffList[i].name}.pdf");

                    using (var sw = new StreamWriter(outputStaffInvoiceHtmlFolder + $@"{allStaffList[i].name}.txt"))
                    {
                        sw.WriteLine(html);
                    }
                }
            }
            return "";

        }

        public static string companyHistory(List<company> companyList)
        {
            Workbook wb = new Workbook();

            // 得到第一個工作表。
            Worksheet sheet1 = wb.Worksheets[0];
            sheet1.Cells.SetColumnWidth(1, 20.0);
            sheet1.Cells.SetColumnWidth(2, 30.0);
            sheet1.Cells.SetColumnWidth(3, 10.0);
            sheet1.Cells.SetColumnWidth(4, 10.0);
            sheet1.Cells.SetColumnWidth(5, 10.0);
            sheet1.Cells.SetColumnWidth(6, 10.0);
            sheet1.Cells.SetColumnWidth(7, 10.0);
            // 獲取工作表的單元格集合
            Cells cells = sheet1.Cells;

            // 為單元格設置值
            Aspose.Cells.Cell cell = cells["A1"];
            cell.PutValue("院舍名稱");
            cell = cells["B1"];
            cell.PutValue("聯絡人");
            cell = cells["C1"];
            cell.PutValue("地址");
            cell = cells["D1"];
            cell.PutValue("發票號碼");
            cell = cells["E1"];
            cell.PutValue("發票金額");
            for (int i = 0; i < companyList.Count; i++)
            {
                cell = cells[@$"A{i + 2}"];
                cell.PutValue(companyList[i].customerName);
                cell = cells[@$"B{i + 2}"];
                cell.PutValue(companyList[i].contactPeople);
                cell = cells[@$"C{i + 2}"];
                cell.PutValue(companyList[i].address);
                cell = cells[@$"D{i + 2}"];
                cell.PutValue(companyList[i].invoiceNum);
                cell = cells[@$"E{i + 2}"];
                cell.PutValue(companyList[i].invoiceTotalSalary);
               
            }

            // 保存 Excel 文件。
            wb.Save(outputFolder + "company_output.xlsx", SaveFormat.Xlsx);

            return "";
        }
    }
}

