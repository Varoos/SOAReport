using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    
    public class LeaseRevList
    {
        public string Property { get; set; }
        public string Unit { get; set; }
        public string TCNo { get; set; }
        public string UnitType { get; set; }
        public string Usage { get; set; }
        public decimal ContractValue { get; set; }
        public string Status { get; set; }
        public string Tenant { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string PostedDate { get; set; }
        public string TerminationDate { get; set; }
        public int TotContractDays { get; set; }
        public decimal DayRent { get; set; }
        public int TotRevDays { get; set; }
        public decimal Am_Amt { get; set; }
        public decimal DeferredVal { get; set; }
        public decimal PDC { get; set; }
        public decimal Security { get; set; }
        public decimal OtherIncome { get; set; }
    }
    
    public class LeaseRevCls
    {
        public LostRevFilter _filter { get; set; }
        public List<LeaseRevList> _list { get; set; }
    }
    
}