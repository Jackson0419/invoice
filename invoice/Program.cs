
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
    table,
    td,
    th {border:1px solid;
    }* {
 font-size: 100%;
 font-family: Times New Roman, Times, serif;
}
 
    table {width: 100%;
      border-collapse: collapse;
    
            }

    td {font - weight: bold;
    }

    .center {margin: auto;
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
        public static List<companyDuty> companyDuties = new List<companyDuty>();
 


        static async Task Main(string[] args)
        {
           
            try
            {  
              
                Console.OutputEncoding = Encoding.Unicode;

                string url = $"http://58.176.128.146/test.php";



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


                Console.WriteLine("將 timesheet.xlsx 放入資料夾 " + timeSheetFolder);
                Console.WriteLine("將 logo.jpg和company_salary.xlsx和staff_salary.xlsx和bank.xlsx 放入資料夾 " + generalFolder);
                Console.WriteLine("完成後請輸入 1 然後按 ENTER");
                string check1 = Console.ReadLine();
                if (check1 == "1")
                {





                    using var client = new HttpClient();

                    List<string> JobIdList = new List<string>();
                    var response = client.GetAsync(url).GetAwaiter().GetResult();




                   
                     var content = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();// response value
                    //string content = "success";
                    if (content == "success")
                    {

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
                            List<specialEvent> specialEventsList = new List<specialEvent>();
                            List<staffList> staffNameList = new List<staffList>();
                            company company = new company();

                            Worksheet worksheet = wb.Worksheets[v];

                            // 打印工作表名稱
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
                                    if(worksheet.Cells[6, j].Value.Equals("RUSH FEE"))
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

                                            if(name != null)
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

                                                    var ExeclRowStaffName = worksheet.Cells[i, j].Value.ToString();

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


                            var staffList = staffNameList;

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

                            for (int i = 0; i < staffList.Count; i++)
                            {

                                for (int q = 0; q < staffList[i].duty.Count; q++)
                                {


                                    for (int e = 0; e < companySalaryList.Count; e++)
                                    {
                                        if (companySalaryList[e].hours == staffList[i].duty[q].dutyHours && companySalaryList[e].title == staffList[i].duty[q].title)
                                        {


                                            staffList[i].duty[q].companySalary = companySalaryList[e].salary;

                                        }

                                    }
                                }
                            }
                            for (int i = 0; i < staffList.Count; i++)
                            {

                                for (int q = 0; q < staffList[i].duty.Count; q++)
                                {


                                    for (int e = 0; e < salaryList.Count; e++)
                                    {
                                        if (salaryList[e].hours == staffList[i].duty[q].dutyHours && salaryList[e].title == staffList[i].duty[q].title)
                                        {

                                            
                                            staffList[i].duty[q].salary = salaryList[e].salary;
                                             
                                        }

                                    }
                                }
                            }



                            for (int i = 0; i < staffList.Count; i++)
                            {

                                staffList[i].duty = staffList[i].duty.OrderBy(x => x.date).ToList();

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
                                        System.Diagnostics.Debug.WriteLine(BankSheet.Cells[e, i].Value.ToString());
                                        for (int q = 0; q < staffList.Count; q++)
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



                            for (int i = 0; i < staffNameList.Count; i++)
                            {
                                if (i == 10)
                                {
                                    var ggw = "";
                                }
                                for (int e = 0; e < staffNameList[i].duty.Count; e++)
                                {
                                    var salary = Convert.ToInt32(staffNameList[i].duty[e].salary);
                                    var companySalary = Convert.ToInt32(staffNameList[i].duty[e].companySalary);
                                    for (int p = 0; p < specialEventsList.Count; p++)
                                    {
                                        if (specialEventsList[p].name == staffList[i].name && specialEventsList[p].date == staffList[i].duty[e].date && specialEventsList[p].shift == staffList[i].duty[e].dutyTime)
                                        {

                                            if (specialEventsList[p].salary == "T8")
                                            {
                                                salary *= 2;
                                                companySalary *= 2;
                                                staffList[i].duty[e].salary = salary.ToString();
                                                staffList[i].duty[e].companySalary = companySalary.ToString();
                                            }
                                            else
                                            {
                                                decimal minutes = Convert.ToInt32(staffList[i].duty[e].dutyHours) * 60;
                                                decimal revisedSalary = (minutes + Convert.ToInt32(specialEventsList[p].salary))/ minutes * salary;
                                                decimal revisedCompanySalary = (minutes + Convert.ToInt32(specialEventsList[p].salary)) / minutes * companySalary;
                                                salary = decimal.ToInt32(revisedSalary);
                                                companySalary = decimal.ToInt32(revisedCompanySalary);
                                                staffList[i].duty[e].salary = decimal.ToInt32(revisedSalary).ToString();
                                                staffList[i].duty[e].companySalary = decimal.ToInt32(revisedCompanySalary).ToString();

                                            }
                                        }
                                    }
                                    staffNameList[i].totalSalaryForCompany += companySalary;
                                    staffNameList[i].totalSalary += salary;

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
                                    duty = staffNameList[i].duty
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

                        await allStaffInvoice(allStaffList);
                        bankAccount(allStaffList);
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
                    List<string> titleList = new List<string>();
                    for (int j = 0; j < companyList[q].staffLists[i].duty.Count; j++)
                    {
                        titleList.Add(companyList[q].staffLists[i].duty[j].title);
                    }
                    titleList = titleList.Distinct().ToList();
                    string title = string.Join(",", titleList);
                    string description = string.Empty;
                    string staffDescription = string.Empty;

                    var dateList = companyList[q].staffLists[i].duty.Select(x => new { x.date, x.shift }).ToList();

                    for (int e = 0; e < dateList.Count; e++)
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

                    }
                    companyList[q].staffLists[i].title = title;
                    companyList[q].staffLists[i].pdfDescription = staffDescription;

                    body += @$"<tr>
               <td style= 'font-family: verdana'>{companyList[q].staffLists[i].name}</td>
               <td>{title}</td>
               <td>{description}</td>
               <td style='text-align: center;'>{companyList[q].staffLists[i].duty.Count}</td>
               <td style='text-align: right;'>{companyList[q].staffLists[i].totalSalaryForCompany}</td>
             </tr>";
                }

                var pdf = await renderer.RenderHtmlAsPdfAsync($@" <!DOCTYPE html>
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
                  <td>{companyList[q].customerName} <br> {companyList[q].address} <br>{companyList[q].contactPeople} </td>
                  <td>Date <div style='float: right;'>{companyList[q].invoiceDate}</div>
                    <br>Invoice No. <div style='float: right;'>{companyList[q].invoiceNum + companyList[q].invoiceMonth}</div>
                   
                  </td>
                </tr>
              </table>
              <table>
                <tr>
                  <td>Name</td>
                  <td>Title</td>
                  <td>Description</td>
                  <td>Day</td>
                  <td>Total Amount HK$</td>
                </tr>
                {body}
                <tr>
                  <td colspan=""4"" style='text-align: right;'>Total HK$</td>
                  <td style='text-align: right;'>{allTotal}</td>
                </tr>
              </table>
              <table>
                <tr>
                  <td style='font-size: 11px;font-family: verdana'>This payment is now due. Please settle the payment as soon as possible. Cheque should be payable to ‘Hygiene First Company Limited’ <br>
                    <br> The amount may also be directly deposited into our BOC account : (Account Number: 012-742-2-019880-4). <br>
                    <br>Please email info@hygienefirstgroup.com / Whatsapp 6086 2287 the bank slip to us for our checking. For billing enquiries, please contact our Accounting Department at 3618 9330 (Mr Chau).
                  </td>
                </tr>
                <tr>
                  <td style='font-size: 11px;font-family: verdana'> Late Payment Surcharge : <br> Should bills remain unpaid after 7 days after the postmark date on the envelope, a 5% surcharge will be added to the outstanding amount. <br> After 14 days, a 10% interest will be imposed on the outstanding amount. </td>
                </tr>
              </table>
              <br>
              <b>For and on behalf of <br> Hygiene First Company Limited </b>
            </div>
          </body>
        </html>");

                pdf.SaveAs(companyList[q].companyOutPutPath + companyList[q].customerName + companyList[q].invoiceNoForCompanySalary + "總invoice_"+ companyList[q].invoiceMonth +"月"+ ".pdf");

            }



            //--------------------------------Single

            return "123";
        }

        public static async Task<string> allStaffInvoice(List<allStaff> allStaffList)
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
                    body += @$"<tr>
                 <td width=""20%"">{allStaffList[i].companyDuty[q].companyName}
                  <td width=""10%"">{allStaffList[i].name}</td>
                  <td width=""10%"">{allStaffList[i].companyDuty[q].title}</td>
                  <td>{allStaffList[i].companyDuty[q].pdfDescription}</td>
                <td style = 'text-align: center;'>{allStaffList[i].companyDuty[q].duty.Count}</td>
                <td style = 'text-align: right;'>{allStaffList[i].companyDuty[q].oldTotalSalary}</td>
                </tr> 
               ";




                    totalSalary += allStaffList[i].companyDuty[q].oldTotalSalary;
                }
                if (allStaffList[i].cancelFees != null)
                {
                    totalSalary += decimal.Parse(allStaffList[i].cancelFees);
                  

                    body += @$"<tr>
       <td>N/A</td>
                   <td>{allStaffList[i].name}</td>
                   <td>N/A</td>
                   <td>取消費</td>
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
                   <td style='text-align: center;'>N/A</td>
                   <td style='text-align: right;'>{allStaffList[i].otherFees}</td>
                 </tr>";

                }

               




                allStaffList[i].totalSalary = totalSalary.ToString();

                var pdf = await renderer.RenderHtmlAsPdfAsync($@" <!DOCTYPE html>
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
              <table>
                <tr>
                  <td>Company</td>
                  <td>Name</td>
                  <td>Title</td>
                  <td width=""40%"">Description</td>
                  <td>Day</td>
                  <td width=""25%"">Total Amount HK$</td>
                </tr>
                {body}
                <tr>
                  <td colspan=""5"" style='text-align: right;'>Total HK$</td>
                  <td style='text-align: right;'>{totalSalary}</td>
                </tr>
              </table>
              <table>
                <tr>
                  <td style='font-size: 11px;font-family: verdana'>This payment is now due. Please settle the payment as soon as possible. Cheque should be payable to ‘Hygiene First Company Limited’ <br>
                    <br> The amount may also be directly deposited into our BOC account : (Account Number: 012-742-2-019880-4). <br>
                    <br>Please email info@hygienefirstgroup.com / Whatsapp 6086 2287 the bank slip to us for our checking. For billing enquiries, please contact our Accounting Department at 3618 9330 (Mr Chau).
                  </td>
                </tr>
                <tr>
                  <td style='font-size: 11px;font-family: verdana'> Late Payment Surcharge : <br> Should bills remain unpaid after 7 days after the postmark date on the envelope, a 5% surcharge will be added to the outstanding amount. <br> After 14 days, a 10% interest will be imposed on the outstanding amount. </td>
                </tr>
              </table>
              <br>
              <b>For and on behalf of <br> Hygiene First Company Limited </b>
            </div>
          </body>
        </html>");

                pdf.SaveAs(outputFolder + "\\staff_invoice\\" + $@"{allStaffList[i].name}.pdf");

            }
            return "";
         
        }
    }
}

