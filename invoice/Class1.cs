using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace invoice
{
    public class company
    {
        public string customerName { get; set; }
        public string address { get; set; }
        public string contactPeople { get; set; }
        public string month { get; set; }
        public string invoiceDate { get; set; }
        public string invoiceNum { get; set; }
        public string invoiceNoForCompanySalary { get; set; }
        public string invoiceMonth { get; set; }
        public string companyOutPutPath { get; set; }
        public string receiptDate { get; set; }
        public string invoiceTotalSalary { get; set; }

        public List<staffList> staffLists = new List<staffList>();
    }

    public class companyDuty
    {
        public List<tittleDuty> titleDuty = new List<tittleDuty>();
        public List<duty> duty = new List<duty>();
        public string pdfDescription { get; set; }
        public string companyName { get; set; }
        public string title { get; set; }
        public decimal oldTotalSalary { get; set; }

    }

    public class staffList
    {
        public string name { get; set; }
        public string engName { get; set; }
        public string bankAccount { get; set; }
        public string firstRegisterFees { get; set; }
        public string uniformFees { get; set; }
        public string cancelFees { get; set; }
        public string otherFees { get; set; }
        public string urgentFees { get; set; }
        public string bonus { get; set; }
        public string transportFees { get; set; }
        public string remark { get; set; }
        public decimal totalStaffSalary { get; set; }
        public decimal totalSalaryOld { get; set; }
        public decimal totalSalaryForCompany { get; set; }
        public string pdfDescription { get; set; }
        public string title { get; set; }
        public List<duty> duty = new List<duty>();

        public List<tittleDuty> titleDuty = new List<tittleDuty>();

    }
    public class allStaff
    {
        public string name { get; set; }
        public string engName { get; set; }
        public string bankAccount { get; set; }
        public string firstRegisterFees { get; set; }
        public string uniformFees { get; set; }
        public string cancelFees { get; set; }
        public string otherFees { get; set; }
        public string urgentFees { get; set; }
        public string bonus { get; set; }
        public string transportFees { get; set; }
        public string remark { get; set; }
        public string totalSalary { get; set; }
        public List<companyDuty> companyDuty = new List<companyDuty>();
    }
    public class duty
    {
        public string date { get; set; }
        public string dutyTime { get; set; }
        public string shift { get; set; }
        public string dutyHours { get; set; }
        public string title { get; set; }
        public string salary { get; set; }
        public string companySalary { get; set; }
        public string OTreason { get; set; }
        public string T8reason { get; set; }
        public string companyUnitPrice { get; set; }
        public string staffUnitPrice { get; set; }

        public double T8StaffRemovesalary { get; set; }
        public double T8CompanyRemoveSalary { get; set; }
        public double T8StaffAddsalary { get; set; }
        public double T8CompanyAddSalary { get; set; }
        public string T8CompanySalaryFormula { get; set; }
        public string T8StaffSalaryFormula { get; set; }
        public decimal t8CompanySalary { get; set; }
        public decimal t8StaffSalary { get; set; }
        public bool T8 { get; set; }
        public bool OT { get; set; }
        public bool bonus { get; set; }
        public decimal bonusSalary { get; set; }
        public string BonusReason { get; set; }
        public decimal OTcompanySalary { get; set; }
        public decimal OTStaffsalary { get; set; }
        public string OTCompanySalaryFormula { get; set; }
        public string OTStaffSalaryFormula { get; set; }
    }

    public class hourSalary
    {
        public string hours { get; set; }
        public string title { get; set; }
        public string salary { get; set; }
    }

    public class specialEvent
    {
        public string date { get; set; }
        public string name { get; set; }
        public string shift { get; set; }
        public string hours { get; set; }
        public string eventT8orOT { get; set; }
        public string reason { get; set; }
    }

    public class tittleDuty {
        public string title { get; set; }
        public string hours { get; set; }
        public string companyUnitPrice { get; set; }
        public string staffUnitPrice { get; set; }
        public List<duty> dutyList { get; set; }


        }

    public class totalAmountObj
    {
        public decimal totalAmount { get; set; }
        public List<eachTotal> eachTotal = new List<eachTotal>();
    }
    public class eachTotal
    {
        public decimal total { get; set; }
        public string name { get; set; }
    }

    public class titleTotalAmount
    {
        public string title { get; set; }
        public List<titleTotalAmountObj> companys { get; set; }
        public List<titleTotalAmountObj> staffs { get; set; }
    }
    public class titleTotalAmountObj
    {
        public string name { get; set; }
        public decimal amount { get; set; }
    }
}
