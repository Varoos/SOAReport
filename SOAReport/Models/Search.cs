using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    public class Search
    {
        public int CompanyId { get; set; }
        public string ReportDate { get; set; }
        public string TC_No { get; set; }
        public string TC_Code { get; set; }
        public int TC_No_Id { get; set; }
        public string Tenant { get; set; }
        public int TenantId { get; set; }
        public int PropertyId { get; set; }
        public string PropertyName { get; set; }
        public int UnitId { get; set; }
        public string UnitName { get; set; }
        public string TerminationDate { get; set; }
        public string ContractStartDate { get; set; }
        public string ContractEndDate { get; set; }
        public decimal ContractAmount { get; set; }
        public int AccountId { get; set; }
        public string AccountName { get; set; }
    }
}