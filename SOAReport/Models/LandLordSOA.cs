using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    public class LandLordSOA
    {
        public LLSOA_Header _header { get; set; }
        public List<LLSOA_List> _listData { get; set; }
    }
    public class LLSOA_Header
    {
        public string Tenant { get; set; }
        public string TenantCode { get; set; }
        public string Address { get; set; }
        public string Mobile { get; set; }
        public string Email { get; set; }
        public string ReportDate { get; set; }
        public decimal SecDep { get; set; }
        public decimal OpBal { get; set; }
        public decimal SecDep2 { get; set; }
        public decimal OpBal2 { get; set; }
        public decimal TotalDebit { get; set; }
        public decimal TotalCredit { get; set; }
        public decimal RentBal { get; set; }
        public decimal ContFund { get; set; }
        public decimal AccBal { get; set; }
        public decimal RentBal2 { get; set; }
        public decimal ContFund2 { get; set; }
        public decimal AccBal2 { get; set; }
        public decimal CompanyId { get; set; }
    }
    public class LLSOA_List
    {
        public string DocDate { get; set; }
        public string DocNo { get; set; }
        public string Desc { get; set; }
        public string Account { get; set; }
        public decimal Debit { get; set; }
        public decimal Credit { get; set; }
        public decimal balance { get; set; }
        public decimal balance2 { get; set; }
    }

}