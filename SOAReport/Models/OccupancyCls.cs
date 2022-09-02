using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    public class OccupancyFilter
    {
        public string PropertyGrp { get; set; }
        public string Property { get; set; }
        public string Unit { get; set; }
        public string TCNoid { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string UnitType { get; set; }
        public string Usage { get; set; }
        public decimal Sqft { get; set; }
        public decimal SqftTo { get; set; }
        public string PropertyGrpName { get; set; }
        public string PropertyName { get; set; }
        public string UnitName { get; set; }
        public string TcNos { get; set; }
        public string UnitTypeName { get; set; }
        public string UsageName { get; set; }
        public string CompanyId { get; set; }
    }
    public class OccupancyList
    {
        public string PropertyGrp { get; set; }
        public string Property { get; set; }
        public string Unit { get; set; }
        public string TCNo { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string UnitType { get; set; }
        public string Usage { get; set; }
        public decimal Sqft { get; set; }
        public int AccPrdDays { get; set; }
        public decimal ContractAmt { get; set; }
        public string AmFrom { get; set; }
        public string AmTo { get; set; }
        public int AmDays { get; set; }
        public decimal AmAmt { get; set; }
        public string Status { get; set; }
        public decimal dayRent { get; set; }
        public decimal sqRent { get; set; }
        public int vacdays { get; set; }
        public decimal AnnualRent { get; set; }
        public decimal vacLoss { get; set; }
        public decimal UnAm_Amt { get; set; }
        public int StayedDays { get; set; }
        public string TCStartDate { get; set; }
        public string TCEndDate { get; set; }
        public decimal StayedAmortizedValue { get; set; }
        public decimal TotalAmortizedValue { get; set; }
        public decimal Unearnedremaining { get; set; }
    }
    public class OccupancyCls
    {
        public OccupancyFilter _filter { get; set; }
        public List<OccupancyList> _list { get; set; }
    }
}