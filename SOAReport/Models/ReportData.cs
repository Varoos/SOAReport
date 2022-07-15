using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    public class ReportData
    {
        public Search _searchObj { get; set; }
        public List<ListData> _listData { get; set; }
        public List<AgeingData> _ageingData { get; set; }
    }
    public class SelectOptions
    {
        public string Name { get; set; }
        public string Code { get; set; }
        public string Id { get; set; }
        public string FId { get; set; }
        public string Extra { get; set; }
    }
    public class ListData
    {
        public string DocDate { get; set; }
        public string DocNo { get; set; }
        public string Account { get; set; }
        public decimal Debit { get; set; }
        public decimal Credit { get; set; }
        public decimal balance { get; set; }
        public string PDCNo { get; set; }
        public string sChequeNo { get; set; }
        public int No_ofdays_stayed { get; set; }
        public int pdccount { get; set; }
    }

    public class AgeingData
    {
        public string DocDate { get; set; }
        public string DocNo { get; set; }
        public string Tenant_Name { get; set; }
        public string Tenant_Code { get; set; }
        public string TC_No { get; set; }
        public decimal Balance { get; set; }
        public decimal Days30 { get; set; }
        public decimal Days60 { get; set; }
        public decimal Days90 { get; set; }
        public decimal Days120 { get; set; }
        public decimal Days150 { get; set; }
        public decimal Days180 { get; set; }
        public decimal Days360 { get; set; }
        public decimal Days_360 { get; set; }
        public string PDCNo { get; set; }
        public string sChequeNo { get; set; }
        public int No_ofdays_stayed { get; set; }
        public int pdccount { get; set; }
        public string Balance2 { get; set; }
        public string Days302 { get; set; }
        public string Days602 { get; set; }
        public string Days902 { get; set; }
        public string Days1202 { get; set; }
        public string Days1502 { get; set; }
        public string Days1802 { get; set; }
        public string Days3602 { get; set; }
        public string Days_3602 { get; set; }
    }
}