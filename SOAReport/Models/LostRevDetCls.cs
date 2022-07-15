using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    public class LostRevFilter
    {
        public string PropertyGrp { get; set; }
        public string Property { get; set; }
        public string Unit { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string PropertyGrpName { get; set; }
        public string PropertyName { get; set; }
        public string UnitName { get; set; }
        public string CompanyId { get; set; }
    }
    public class LostRevDetList
    {
        public string Property { get; set; }
        public string Unit { get; set; }
        public decimal AnnualRent { get; set; }
        public string LostStartDt { get; set; }
        public string LostEndDt { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public int TotLostDays { get; set; }
        public decimal TotLostRev { get; set; }
        public decimal Jan { get; set; }
        public decimal Feb { get; set; }
        public decimal Mar { get; set; }
        public decimal Apr { get; set; }
        public decimal May { get; set; }
        public decimal Jun { get; set; }
        public decimal Jul { get; set; }
        public decimal Aug { get; set; }
        public decimal Sep { get; set; }
        public decimal Oct { get; set; }
        public decimal Nov { get; set; }
        public decimal Dec { get; set; }
    }
    public class LostRevDetCls
    {
        public LostRevFilter _filter { get; set; }
        public List<LostRevDetList> _list { get; set; }
    }
    public class LostRevCls
    {
        public LostRevFilter _filter { get; set; }
        public List<LostRevList> _list { get; set; }
    }
    public class LostRevList
    {
        public string Property { get; set; }
        public string Unit { get; set; }
        public decimal AnnualRent { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public int TotLostDays { get; set; }
        public decimal TotLostRev { get; set; }
    }
}