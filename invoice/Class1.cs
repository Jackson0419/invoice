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
        public string invoiceCustomerNum { get; set; }
        public string invoiceMonth { get; set; }
        public string companyOutPutPath { get; set; }


        public List<staffList> staffLists = new List<staffList>();
    }
  
    public class companyDuty
    { 
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
        public string remark { get; set; }
        public decimal totalSalary { get; set; }
        public decimal totalSalaryOld { get; set; }
        public string pdfDescription { get; set; }
        public string title { get; set; }
        public List<duty> duty = new List<duty>();

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



    }

    public class hourSalary
    {
        public string hours { get; set; }
        public string title { get; set; }
        public string salary { get; set; }
    }
}
