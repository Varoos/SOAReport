using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    public class ExpFilter
    {
        public string PropertyGrp { get; set; }
        public string Property { get; set; }
        public string Unit { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string UnitType { get; set; }
        public string PropertyGrpName { get; set; }
        public string PropertyName { get; set; }
        public string UnitName { get; set; }
        public string UnitTypeName { get; set; }
        public string CompanyId { get; set; }
    }
    public class ExpList
    {
        public string PropertyGrp { get; set; }
        public string Property { get; set; }
        public string Unit { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string UnitType { get; set; }
        public string Usage { get; set; }
        public string Sqft { get; set; }
        public int AccPrdDays { get; set; }
        public decimal Amt { get; set; }
        public string Account { get; set; }
    }
    public class ExpenseCls
    {
        public ExpFilter _filter { get; set; }
        public List<ExpList> _list { get; set; }
    }
}