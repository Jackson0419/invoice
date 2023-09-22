
using Aspose.Cells;
using PugPdf.Core;
using System;
using System.Collections.Generic;
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
     table-layout: fixed;
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
        public static string outputStaffInvoiceHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\ouputHtml\\staff\\";
        public static string inputHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\inputHtml\\";
        public static string outputHtmlFolder = desktopPath + "\\" + "invoice\\" + DateTime.Now.ToString("yyyyMMdd") + "\\inputHtml\\output\\";
     
        public static List<companyDuty> companyDuties = new List<companyDuty>();
        public static string invoiceDate=string.Empty;


        static async Task Main(string[] args)
        {

            try
            {

                Console.OutputEncoding = Encoding.Unicode;

                string url = $"https://testsds123-669967cd5270.herokuapp.com/";



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
                
                 


                using var client = new HttpClient();
                var response = client.GetAsync(url).GetAwaiter().GetResult();
                var content = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();// response value
                Console.WriteLine("將 timesheet.xlsx 放入資料夾 " + timeSheetFolder);
                Console.WriteLine("將 logo.jpg和company_salary.xlsx和staff_salary.xlsx和bank.xlsx 放入資料夾 " + generalFolder);
                Console.WriteLine("1 = Gen Invoice, 2 = Gen HtmlCode");
                string check1 = Console.ReadLine();
                if (content == "success")
                {

                    if(check1 == "2")
                    {
                        var renderer = new HtmlToPdf();
                        DirectoryInfo d = new DirectoryInfo(inputHtmlFolder);
                        var matchFolder = d.GetFiles("*.txt");
                        for (int i =0;i< matchFolder.Length;i++)  
                        {
                            Console.WriteLine(matchFolder[i].FullName +" Processing");
                            string text = File.ReadAllText(matchFolder[i].FullName);  
                            var pdf = await renderer.RenderHtmlAsPdfAsync(text);
                            var ppp = Path.GetFileNameWithoutExtension(matchFolder[i].FullName);
                            pdf.SaveAs(outputHtmlFolder + Path.GetFileNameWithoutExtension(matchFolder[i].FullName) + ".pdf");
                            Console.WriteLine(matchFolder[i].FullName + " Done");
                        }

                                               
              
                    }
                    if (check1 == "1")
                {





               

                    List<string> JobIdList = new List<string>();
                    
                    //string content = "success";
                    

                        Workbook wb = new Workbook(timeSheetFolder + "timesheet.xlsx");
                        /*
                                                for (int q = 0; q < wb.Worksheets.Count(); q++)
                                                {
                                                    Worksheet worksheetTest = wb.Worksheets[q];
                                                    companyDuties.Add
                                                        (new companyDuty
                                                        {
                                                            companyName = worksheetTest.Cells[3, 4].Value != null ? worksheetTest.Cells[3, 4].Value.ToString() : null,
                                                            duty = new List<duty>()

                                                        });

                                                }
                        */

                        // 使用其索引獲取工作表

                        for (int v = 0; v < wb.Worksheets.Count(); v++)
                        {
                            Console.WriteLine(v);
                            List<specialEvent> specialEventsList = new List<specialEvent>();
                            List<staffList> staffNameList = new List<staffList>();
                            company company = new company();

                            Worksheet worksheet = wb.Worksheets[v];
                            var ggg = worksheet.IsVisible;
                            // 打印工作表名稱
                            if (worksheet.IsVisible == false)
                            {
                                continue;
                            }
                            
                            Console.WriteLine("Worksheet: " + worksheet.Name);

                            // 獲取行數和列數
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

                                            var date = worksheet.Cells[i, j].Value != null ? DateTime.Parse(worksheet.Cells[i, j].Value.ToString()).ToString("dd-MM-yyyy") : null;
                                            var name = worksheet.Cells[i, j + 1].Value != null ? worksheet.Cells[i, j + 1].Value.ToString() : null;
                                            var shift = worksheet.Cells[i, j + 2].Value != null ? worksheet.Cells[i, j + 2].Value.ToString() : null;
                                            var salary = worksheet.Cells[i, j + 3].Value != null ? worksheet.Cells[i, j + 3].Value.ToString() : null;
                                            var reason = worksheet.Cells[i, j + 4].Value != null ? worksheet.Cells[i, j + 4].Value.ToString() : null;

                                            if (name != null)
                                            {
                                                specialEventsList.Add(new specialEvent
                                                {
                                                    date = date,
                                                    name = name,
                                                    shift = shift,
                                                    salary = salary,
                                                    reason = reason
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
                                            System.Diagnostics.Debug.WriteLine(BankSheet.Cells[e, i].Value.ToString());
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
                                                    staffNameList[q].remark = BankSheet.Cells[e, 8].Value != null ? BankSheet.Cells[e, 8].Value.ToString() : null;


                                                }

                                            }
                                        }

                                    }

                                }
                            }

                        


                            for (int i = 0; i < staffNameList.Count; i++)
                            {
                                if (i == 10)
                                {
                                    var ggw = "";
                                }
                                for (int e = 0; e < staffNameList[i].duty.Count; e++)
                                {
                                    var salary = Convert.ToDouble(staffNameList[i].duty[e].salary);
                                    var companySalary = Convert.ToDouble(staffNameList[i].duty[e].companySalary);
                                    for (int p = 0; p < specialEventsList.Count; p++)
                                    {
                                        if (specialEventsList[p].name == staffNameList[i].name && specialEventsList[p].date == staffNameList[i].duty[e].date && specialEventsList[p].shift == staffNameList[i].duty[e].dutyTime)
                                        {

                                            staffNameList[i].duty[e].reason = specialEventsList[p].reason;

                                            if (specialEventsList[p].salary == "T8")
                                            {
                                                salary *= 1.5;
                                                companySalary *= 2;
                                                staffNameList[i].duty[e].salary = salary.ToString();
                                                staffNameList[i].duty[e].companySalary = companySalary.ToString();
                                            }
                                            else
                                            {

                                                decimal minutes = Convert.ToDecimal(staffNameList[i].duty[e].dutyHours) * 60;
                                                double revisedSalary = (Convert.ToDouble(minutes) + Convert.ToDouble(specialEventsList[p].salary)) / Convert.ToDouble(minutes) * salary;
                                                decimal revisedCompanySalary = (minutes + Convert.ToInt32(specialEventsList[p].salary)) / minutes * Convert.ToInt32(companySalary);
                                                salary = Convert.ToInt32(revisedSalary);
                                                companySalary = decimal.ToInt32(revisedCompanySalary);
                                                staffNameList[i].duty[e].salary = Convert.ToInt32(revisedSalary).ToString();
                                                staffNameList[i].duty[e].companySalary = decimal.ToInt32(revisedCompanySalary).ToString();

                                            }
                                        }
                                    }
                                    staffNameList[i].totalSalaryForCompany += Convert.ToDecimal(companySalary);
                                    staffNameList[i].totalSalary += Convert.ToDecimal(salary);

                                }

                                staffNameList[i].totalSalaryOld = staffNameList[i].totalSalary;

                                if (staffNameList[i].firstRegisterFees != null)
                                {

                                    staffNameList[i].totalSalary += Convert.ToDecimal(staffNameList[i].firstRegisterFees);

                                }
                                if (staffNameList[i].uniformFees != null)
                                {


                                    staffNameList[i].totalSalary += Convert.ToDecimal(staffNameList[i].uniformFees);

                                }
                                if (staffNameList[i].cancelFees != null)
                                {


                                    staffNameList[i].totalSalary += Convert.ToDecimal(staffNameList[i].cancelFees);

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
                                        uniformFees = staffNameList[i].uniformFees
                                    };
                                    staff.companyDuty.Add(companyDuty);
                                    allStaffList.Add(staff);

                                }
                            }

                            companyList.Add(company);
                        }



                        await company_invoice(companyList);
                        Console.WriteLine("company invoice processing");
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
                        Console.WriteLine("company invoice processing Done");
                        Console.WriteLine("要staff invoice 輸入 1, 不要staff invoice 輸入 2");
                        string check2 = Console.ReadLine();
                        if (check2 == "1" || check2 =="2")
                        {
                            await allStaffInvoice(allStaffList, check2); 
                            bankAccount(allStaffList);
                            companyHistory(companyList);
                        }
                        else
                        {
                            
                        }
                     
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


        public static string bankAccount(List<allStaff> allstaffList)
        {
            Workbook yoyoyo = new Workbook();

            // 得到第一個工作表。
            Worksheet sheet1 = yoyoyo.Worksheets[0];
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
            cell.PutValue("TOTAL");
            cell = cells["I1"];
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
                cell.PutValue(allstaffList[i].totalSalary);
                cell = cells[@$"I{i + 2}"];
                cell.PutValue(allstaffList[i].remark);
            }

            // 保存 Excel 文件。
            yoyoyo.Save(outputFolder + "System Bank Record_Program_output.xlsx", SaveFormat.Xlsx);

            return "";
        }
        public static async Task<string> company_invoice(List<company> companyList)
        {
            var renderer = new HtmlToPdf();

            ;
            for (int q = 0; q < companyList.Count; q++)
            {
                decimal allTotal = 0;
                renderer.PrintOptions.Title = companyList[q].companyOutPutPath + companyList[q].customerName + "總invoice_" + companyList[q].invoiceMonth + "月";
                string body = string.Empty;

                for (int i = 0; i < companyList[q].staffLists.Count; i++)
                {

                    allTotal += companyList[q].staffLists[i].totalSalaryForCompany;


                    for (int e = 0; e < companyList[q].staffLists[i].titleDuty.Count; e++)
                    {
                        decimal eachDutytotal = 0;
                        string description = string.Empty;
                        string title = string.Empty;
                        string companyUnitprice = companyList[q].staffLists[i].titleDuty[e].companyUnitPrice;
                        title = companyList[q].staffLists[i].titleDuty[e].title;
                        string dutyCount = "0";
                        for (int o = 0; o < companyList[q].staffLists[i].titleDuty[e].dutyList.Count; o++)
                        {
                            description += DateTime.ParseExact(companyList[q].staffLists[i].titleDuty[e].dutyList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + companyList[q].staffLists[i].titleDuty[e].dutyList[o].shift + ")";
                            if (!string.IsNullOrEmpty(companyList[q].staffLists[i].titleDuty[e].dutyList[o].reason))
                            {
                                description += "(" + companyList[q].staffLists[i].titleDuty[e].dutyList[o].reason + ")";
                            }
                            if (o == companyList[q].staffLists[i].titleDuty[e].dutyList.Count - 1)
                            {

                            }
                            else
                            {
                                description += ", ";

                            }

                            eachDutytotal += decimal.Parse(companyList[q].staffLists[i].titleDuty[e].dutyList[o].companySalary);

                        }
                        body += @$"<tr>
                           <td style= 'font-family: verdana'>{companyList[q].staffLists[i].name}</td>
                           <td>{title}</td>
                           <td width=50px>{description}</td>
                           <td style='text-align: right;'>{companyList[q].staffLists[i].titleDuty[e].companyUnitPrice}</td>
                           <td style='text-align: center;'>{companyList[q].staffLists[i].titleDuty[e].dutyList.Count}</td>
                           <td style='text-align: right;'>{eachDutytotal}</td>
                         </tr>";

                    }

                    /* List<string> titleList = new List<string>();
                    for (int j = 0; j < companyList[q].staffLists[i].duty.Count; j++)
                    {
                        titleList.Add(companyList[q].staffLists[i].duty[j].title);
                    }
                    titleList = titleList.Distinct().ToList();
                    string title = string.Join(",", titleList);
                    string description = string.Empty;
                    string staffDescription = string.Empty;

                    var dateList = companyList[q].staffLists[i].duty.Select(x => new { x.date, x.shift }).ToList();
*/
                    /*   for (int e = 0; e < dateList.Count; e++)
                       {

                           if (e >= 1)
                           {
                               description += ", ";
                               staffDescription += ", ";
                           }
                           if (e != 0)
                           {
                               if (e % 6 == 0)
                               {
                                   description += "<br>";

                               }
                               if( e% 4 == 0)
                               {
                                   staffDescription += "<br>";
                               }
                           }
                           staffDescription += DateTime.ParseExact(dateList[e].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + dateList[e].shift + ")"; 
                           description += DateTime.ParseExact(dateList[e].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + dateList[e].shift + ")";

                       }*/
                    //companyList[q].staffLists[i].title = title;
                    // companyList[q].staffLists[i].pdfDescription = staffDescription;


                    /*       body += @$"<tr>
                      <td style= 'font-family: verdana'>{companyList[q].staffLists[i].name}</td>
                      <td>{title}</td>
                      <td>{description}</td>
                      <td style='text-align: center;'>{companyList[q].staffLists[i].duty.Count}</td>
                      <td style='text-align: right;'>{companyList[q].staffLists[i].totalSalaryForCompany}</td>
                    </tr>";*/
                }
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
              <b>Bill To:</b>
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
                  <td width ='50px'>Name</td>
                  <td width ='30px'>Title</td>
                  <td width ='250px'>Description</td>
                  <td width ='40px'>Unit Price HK$</td>
                  <td width ='25px'>Qty</td>
                  <td width ='100px'>Total Amount HK$</td>
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
                var pdf = await renderer.RenderHtmlAsPdfAsync(html);

                pdf.SaveAs(companyList[q].companyOutPutPath + companyList[q].customerName + companyList[q].invoiceNoForCompanySalary + "總invoice_" + companyList[q].invoiceMonth + "月" + ".pdf");

                using (var sw = new StreamWriter(outputCompanyInvoiceHtmlFolder + companyList[q].customerName + companyList[q].invoiceNoForCompanySalary + "總invoice_" + companyList[q].invoiceMonth + "月.txt"))
                {
                    sw.WriteLine(html);
                }

            }



            //--------------------------------Single

            return "123";
        }

        public static async Task<string> allStaffInvoice(List<allStaff> allStaffList, string option)
        {

            for (int i = 0; i < allStaffList.Count; i++)
            {
                decimal totalSalary = 0;
                string body = string.Empty;
                Console.WriteLine(i + 1 + "/" + allStaffList.Count);
                Console.WriteLine(allStaffList[i].name + " Processing");
                var renderer = new HtmlToPdf();

                for (int q = 0; q < allStaffList[i].companyDuty.Count; q++)
                {



                    renderer.PrintOptions.Title = allStaffList[i].name;
                    for (int e = 0; e < allStaffList[i].companyDuty[q].titleDuty.Count; e++)
                    {
                        string description = string.Empty;
                        decimal eachDutytotal = 0;

                        for (int o = 0; o < allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Count; o++)
                        {
                            description += DateTime.ParseExact(allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].date.ToString(), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd/M") + "(" + allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].shift + ")";

                            if (!string.IsNullOrEmpty(allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].reason))
                            {
                                description += "(" + allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].reason + ")";
                            }

                            if (o == allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Count - 1)
                            {

                            }
                            else
                            {
                                description += ", ";

                            }

                            eachDutytotal += decimal.Parse(allStaffList[i].companyDuty[q].titleDuty[e].dutyList[o].salary);
                        }

                        body += @$"<tr>
                         <td>{allStaffList[i].companyDuty[q].companyName}
                          <td>{allStaffList[i].name}</td>
                          <td>{allStaffList[i].companyDuty[q].titleDuty[e].title}</td>
                          <td>{description}</td>
                        <td>{allStaffList[i].companyDuty[q].titleDuty[e].staffUnitPrice}</td>
                        <td style = 'text-align: center;'>{allStaffList[i].companyDuty[q].titleDuty[e].dutyList.Count}</td>
                        <td style = 'text-align: right;'>{eachDutytotal}</td>
                        </tr> 
                       ";




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






                allStaffList[i].totalSalary = totalSalary.ToString();
                if (option == "1")
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
                  <td  width ='499px'><b>備注：</b><br><b>1.自取現金支票：</b> 行政費用HKD$20/次。(地址：屯門震寰路九號好收成工業大廈10樓04室）<br><b>2.郵寄現金支票: </b>行政費用HKD$30/次 (平郵/郵局掛號 另加HKD$16/順豐到付另加HKD$18)，郵寄風險自負。<br><b>3.首次登記:</b>自僱人士獲派首次配對服務後，本公司將收取一次性HKD$50為首次登記費用。<br><b>4.更改服務酬金方式：</b> 行政費用HKD$15/次。           </td>
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
            Workbook yoyoyo = new Workbook();

            // 得到第一個工作表。
            Worksheet sheet1 = yoyoyo.Worksheets[0];
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

            for (int i = 0; i < companyList.Count; i++)
            {
                cell = cells[@$"A{i + 2}"];
                cell.PutValue(companyList[i].customerName);
                cell = cells[@$"B{i + 2}"];
                cell.PutValue(companyList[i].contactPeople);
                cell = cells[@$"C{i + 2}"];
                cell.PutValue(companyList[i].address);

            }

            // 保存 Excel 文件。
            yoyoyo.Save(outputFolder + "company_output.xlsx", SaveFormat.Xlsx);

            return "";
        }
    }
}

