using ClosedXML.Excel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Focus.Common.DataStructs;
using Newtonsoft.Json;
using SOAReport.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml;

namespace SOAReport.Controllers
{
    public class HomeController : Controller
    {
        string errors1 = "";
        public ActionResult Index(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module with CompanyId = " + CompanyId);
                //var Tenants = GetTenants(CompanyId);
                //ViewBag.Tenants = Tenants;
                //DBClass.SetLog("Got Tenants list. Count = " + Tenants.Count().ToString());
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        public IEnumerable<SelectListItem> GetTenants(int cid)
        {
            try
            {
                DBClass.SetLog("GetTenants Event. CompanyId = " + cid.ToString());
                string retrievequery = string.Format(@"exec pCore_CommonSp @Operation=GetTenants");
                DBClass.SetLog("GetTenants Event. Query = " + retrievequery);
                List<SelectListItem> containers = new List<SelectListItem>();
                DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
                DBClass.SetLog("GetTenants Event. ResultSet count = " + ds.Tables[0].Rows.Count.ToString());
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    containers.Add(new SelectListItem()
                    {
                        Text = ds.Tables[0].Rows[i]["sName"].ToString(),//"(" + ds.Tables[0].Rows[i]["sCode"].ToString() + ") " + ds.Tables[0].Rows[i]["sName"].ToString(),
                        Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                    });
                }
                DBClass.SetLog("GetTenants Event. LIst created = " + containers.ToArray());
                return new SelectList(containers.ToArray(), "Value", "Text");
            }
            catch (Exception ex)
            {
                DBClass.SetLog("GetTenants Event. Exception = " + ex.Message);
                return null;
            }
        }
        public ActionResult getTenantslist(int cid, string searchtext)
        {
            string retrievequery = string.Format($@"exec pCore_CommonSp @Operation=SeacrhTenants, @p2='{searchtext}'");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
            string JSONString = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(JSONString, JsonRequestBehavior.AllowGet);
        }
        public JsonResult TenantSelectionChange(int cid, int TenantId)
        {
            try
            {
                DBClass.SetLog("Tenant Selection Change Event. Tenant ID = " + TenantId.ToString());
                string retrievequery = string.Format(@"exec pCore_CommonSp @Operation=GetTcNo_BY_TenantID,@p1=" + TenantId);
                DBClass.SetLog("Tenant Selection Change Event. Query = " + retrievequery);
                List<SelectListItem> containers = new List<SelectListItem>();
                DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
                DBClass.SetLog("Tenant Selection Change Event. ResultSet count = " + ds.Tables[0].Rows.Count.ToString());
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    containers.Add(new SelectListItem()
                    {
                        Text = "(" + ds.Tables[0].Rows[i]["sCode"].ToString() + ") " + ds.Tables[0].Rows[i]["sName"].ToString(),
                        Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                    });
                }
                DBClass.SetLog("Tenant Selection Change Event. LIst created = " + containers.ToArray());
                //return new SelectList(containers.ToArray(), "Value", "Text");
                return Json(containers, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Tenant Selection Change Event. Exception = " + ex.Message);
                return null;
            }
        }
        public DataSet TC_No_SelectionChange(int cid, int TCNo)
        {
            try
            {
                DBClass.SetLog("TC_No Selection Change Event. TC No ID = " + TCNo.ToString());
                string retrievequery = string.Format(@"exec pCore_CommonSp @Operation=Get_TenancyContract_Details,@p1=" + TCNo);
                DBClass.SetLog("TC_No Selection Change Event. Query = " + retrievequery);
                DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
                return ds;
            }
            catch (Exception ex)
            {
                DBClass.SetLog("TC_No Selection Change Event. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public DataSet GetReportData(int cid, string TCNo, string ReportDate)
        {
            try
            {
                DBClass.SetLog("GetReportData. TC No ID = " + TCNo.ToString() + " reportDate = " + ReportDate + " CompanyId = " + cid);
                string retrievequery = string.Format(@"exec Core_SOA_Report_Tenant @TCNo='" + TCNo + "',@ReportDate='" + ReportDate + "'");
                DBClass.SetLog("GetReportData. Query = " + retrievequery);
                DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
                return ds;
            }
            catch (Exception ex)
            {
                DBClass.SetLog("GetReportData. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult Report2(int CompanyId, string ReportDate, int TCNo)
        {
            try
            {
                DBClass.SetLog("Getting Report View. CompanyId = " + CompanyId.ToString() + " ReportDate = " + ReportDate + " TC_No = " + TCNo);
                ReportData _reportData = new ReportData();
                DataSet ds = TC_No_SelectionChange(CompanyId, TCNo);
                if (ds != null)
                {
                    DBClass.SetLog("Getting Report View. TC_No_SelectionChange DataSet is not null");
                    if (ds.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. TC_No_SelectionChange DataSet Table count >0 ");
                        DataRow dr = ds.Tables[0].Rows[0];
                        Search search = new Search();

                        DBClass.SetLog("Getting Report View. GetReportData header data is ready");

                        DataSet ds1 = GetReportData(CompanyId, dr["sCode"].ToString(), ReportDate);
                        if (ds1 != null)
                        {
                            DBClass.SetLog("Getting Report View. GetReportData DataSet is not null");
                            if (ds1.Tables.Count > 0)
                            {
                                DBClass.SetLog("Getting Report View. GetReportData DataSet Table count >0 ");
                                List<ListData> listobj = new List<ListData>();
                                foreach (DataRow dr1 in ds1.Tables[0].Rows)
                                {
                                    listobj.Add(new ListData
                                    {
                                        DocDate = dr1["Date"].ToString(),
                                        DocNo = dr1["Voucher"].ToString(),
                                        Account = dr1["Account"].ToString(),
                                        Debit = Convert.ToDecimal(dr1["Debit"].ToString()),
                                        Credit = Convert.ToDecimal(dr1["Credit"].ToString()),
                                        balance = Convert.ToDecimal(dr1["Balance"].ToString()),
                                        PDCNo = dr1["PDCNo"].ToString(),
                                        sChequeNo = dr1["sChequeNo"].ToString(),
                                        No_ofdays_stayed = Convert.ToInt32(dr1["No_ofdays_stayed"].ToString()),
                                        pdccount = Convert.ToInt32(dr1["pdccount"].ToString()),
                                    });
                                }
                                search.TC_No = dr["sName"].ToString();
                                search.TC_No_Id = TCNo;
                                search.Tenant = dr["Tenant"].ToString();
                                search.PropertyName = dr["Property"].ToString();
                                search.ContractAmount = Convert.ToDecimal(dr["AnnualRent"].ToString());
                                search.UnitName = dr["UnitMaster"].ToString();
                                search.TerminationDate = dr["TerminationDate2"].ToString();
                                search.ContractStartDate = dr["StartDate2"].ToString();
                                search.ContractEndDate = dr["EndDate2"].ToString();
                                search.TC_Code = dr["sCode"].ToString();
                                search.CompanyId = CompanyId;
                                search.ReportDate = ReportDate;
                                _reportData._searchObj = search;
                                _reportData._listData = listobj;
                                DBClass.SetLog("Getting Report View. GetReportData body data is ready");
                            }
                        }
                    }
                }
                TempData["TCNO"] = _reportData._searchObj.TC_Code;
                TempData["ReportDate"] = ReportDate;
                TempData["ReportData"] = _reportData;
                TempData["CompanyId"] = CompanyId;
                TempData.Keep();
                Session["ReportData"] = _reportData;
                if (_reportData == null)
                {
                    return Json("No Data", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    if (_reportData._listData.Count == 0)
                    {
                        return Json("No Data", JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json("Success", JsonRequestBehavior.AllowGet);
                    }
                }
                DBClass.SetLog("Getting Report View. SOA Report body data is ready");
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult Report()
        {
            ReportData _data = (ReportData)Session["ReportData"];
            return View(_data);
        }
        public string GetCompCode(int CompId)
        {
            string CCode = "";
            try
            {
                DBClass.SetLog("GetCompCode CompId= " + CompId.ToString());
                string userid = getDBDetails("User_Id");
                string pwd = getDBDetails("Password");
                string server = getDBDetails("Data_Source");
                DBClass.SetLog("userid = " + userid + ", pwd=" + pwd + ", server=" + server);
                DBClass d = new DBClass(server, userid, pwd);
                string q = "select sCompanyCode from tCore_Company_Details where iCompanyId = " + CompId;
                DBClass.SetLog("q = " + q);
                DataSet dsComp = d.GetData(q);
                CCode = dsComp.Tables[0].Rows[0][0].ToString();
                DBClass.SetLog("CCode = " + CCode);
                return CCode;
            }
            catch (Exception ex)
            {
                DBClass.SetLog("GetCompCode. Exception = " + ex.Message);
                return null;
            }
        }
        public string getDBDetails(string key)
        {
            XmlDocument xmlDoc = new XmlDocument();
            string strFileName = "";
            string sAppPath = AppDomain.CurrentDomain.BaseDirectory;
            strFileName = sAppPath + "\\bin\\XMLFiles\\DBConfig.xml";

            xmlDoc.Load(strFileName);
            XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/DatabaseConfig/Database/" + key + "");
            string strValue;
            XmlNode node = nodeList[0];
            if (node != null)
                strValue = node.InnerText;
            else
                strValue = "";
            return strValue;
        }
        [HttpPost]
        public FileResult ExcelGenerate()
        {

            #region TempData
            ReportData _data = new ReportData();
            _data = (ReportData)TempData["ReportData"];
            Search _head = new Search();
            _head = _data._searchObj;
            TempData.Keep();
            #endregion

            System.Data.DataTable data = new System.Data.DataTable("Statement of Account");
            #region DataColumns
            data.Columns.Add("Date", typeof(string));
            data.Columns.Add("Doc No", typeof(string));
            data.Columns.Add("Account", typeof(string));
            data.Columns.Add("Debit", typeof(decimal));
            data.Columns.Add("Credit", typeof(decimal));
            data.Columns.Add("Balance", typeof(decimal));

            #endregion


            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Statement of Account");
                var dataTable = data;



                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 7));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Statement of Account";

                int r = 3;
                int cell = 2;
                ws.Cell(r, 2).Value = "Tenant";
                ws.Cell(r, 3).Value = _head.Tenant;
                ws.Cell(r, 5).Value = "Date";
                ws.Cell(r, 6).Value = _head.ReportDate;

                r = 4;
                ws.Cell(r, 2).Value = "Property";
                ws.Cell(r, 3).Value = _head.PropertyName;
                ws.Cell(r, 5).Value = "TC No";
                ws.Cell(r, 6).Value = _head.TC_Code;

                r = 5;
                ws.Cell(r, 2).Value = "Unit";
                ws.Cell(r, 3).Value = _head.UnitName;
                ws.Cell(r, 5).Value = "Termination Date";
                ws.Cell(r, 6).Value = _head.TerminationDate;

                r = 6;
                ws.Cell(r, 2).Value = "Contract Start Date";
                ws.Cell(r, 3).Value = _head.ContractStartDate;
                ws.Cell(r, 5).Value = "Contract Amount";
                ws.Cell(r, 6).Value = _head.ContractAmount;

                r = 7;
                ws.Cell(r, 2).Value = "Contract End Date";
                ws.Cell(r, 3).Value = _head.ContractEndDate;
                ws.Cell(r, 5).Value = "No. of Days Stayed";
                ws.Cell(r, 6).Value = _data._listData[0].No_ofdays_stayed;


                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 7));
                //TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(ws.Cell(7, 6), ws.Cell(7, 6)).Style.NumberFormat.Format = "0.00";

                r = 9;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "Date";
                    ws.Cell(r, cell++).Value = "Doc No";
                    ws.Cell(r, cell++).Value = "Account";
                    ws.Cell(r, cell++).Value = "Debit";
                    ws.Cell(r, cell++).Value = "Credit";
                    ws.Cell(r, cell++).Value = "Balance";

                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r, 7));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                int c = 2;

                #region TableLoop
                decimal dr = 0;
                decimal cr = 0;
                decimal bal = 0;
                int count = 1;
                List<ListData> _list = new List<ListData>();
                _list = _data._listData;
                foreach (var obj in _list)
                {
                    dr = Convert.ToDecimal(obj.Debit) + dr;
                    bal = Convert.ToDecimal(obj.balance) + bal;
                    cr = Convert.ToDecimal(obj.Credit) + cr;
                    //count++;
                    c = 2;
                    r++;
                    //ws.Range(ws.Cell(r, c), ws.Cell(r, 7)).Style.Fill.BackgroundColor = XLColor.FromHtml("#90EE90");
                    ws.Cell(r, c++).Value = obj.DocDate;
                    ws.Cell(r, c++).Value = obj.DocNo;
                    ws.Cell(r, c++).Value = obj.Account;
                    ws.Cell(r, c++).Value = obj.Debit.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Credit.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.balance.ToString("N", new CultureInfo("en-US"));
                }


                //Grand Total Row
                r++;
                c = 2;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 4)).Merge().Value = "Total";
                ws.Cell(r, 5).Value = dr.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 6).Value = cr.ToString("N", new CultureInfo("en-US"));
                //ws.Cell(r, 7).Value = bal.ToString("N", new CultureInfo("en-US"));

                ws.Range("B" + r + ":Z" + r + "").Style.Font.Bold = true;
                r++;
                int pdcrow = r;
                if (_data._listData[0].pdccount > 0)
                {
                    ws.Range(ws.Cell(r, 2), ws.Cell(r, 7)).Merge();
                    r = r + 1;
                    wsReportNameHeaderRange = ws.Range(ws.Cell(r++, 2), ws.Cell(r++, 7));
                    wsReportNameHeaderRange.Style.Font.Bold = true;
                    wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.White;
                    wsReportNameHeaderRange.Merge();
                    wsReportNameHeaderRange.Value = "Summary";

                    wsReportNameHeaderRange = ws.Range(ws.Cell(r++, 2), ws.Cell(r++, 7));
                    wsReportNameHeaderRange.Style.Font.Bold = true;
                    wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.White;
                    wsReportNameHeaderRange.Merge();
                    wsReportNameHeaderRange.Value = "Uncleared PDC";
                    cell = 2;
                    ws.Range(ws.Cell(r, 2), ws.Cell(r, 7)).Merge();
                    r = r + 1;
                    #region Headers
                    ws.Range(ws.Cell(r, cell), ws.Cell(r, 3)).Merge().Value = "Doc No";
                    ws.Range(ws.Cell(r, 4), ws.Cell(r, 5)).Merge().Value = "Cheque No";
                    ws.Cell(r, 6).Value = "Cheque Date";
                    ws.Cell(r, 7).Value = "Amount";
                    #endregion
                    foreach (var obj in _list)
                    {
                        if (obj.PDCNo != "0")
                        {
                            c = 2;
                            r++;
                            ws.Range(ws.Cell(r, c), ws.Cell(r, 3)).Merge().Value = obj.DocNo;
                            ws.Range(ws.Cell(r, 4), ws.Cell(r, 5)).Merge().Value = obj.sChequeNo;
                            ws.Cell(r, 6).Value = obj.DocDate;
                            ws.Cell(r, 7).Value = obj.Credit.ToString("N", new CultureInfo("en-US"));
                        }
                    }

                }

                #endregion

                TableRange = ws.Range(ws.Cell(8, 2), ws.Cell(r, 7));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;



                ws.Range(ws.Cell(9, 5), ws.Cell(pdcrow - 1, 7)).Style.NumberFormat.Format = "0.00";
                ws.Cell(7, 6).Style.NumberFormat.Format = "0";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SOAReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        [HttpPost]
        public ActionResult ExportPDF5(int CompanyId, string TcNo, string ReportDt)
        {
            try
            {
                string userid = getDBDetails("User_Id");
                string pwd = getDBDetails("Password");
                string server = getDBDetails("Data_Source");
                string ccode = GetCompCode(Convert.ToInt32(TempData["CompanyId"]));
                string strConnection = "Server=" + server + ";Database=Focus8" + ccode + ";User Id=" + userid + ";Password=" + pwd + ";";
                SqlConnection con = new SqlConnection(strConnection);
                string q = "exec Core_SOA_Report_Tenant @TCNo ='" + TcNo + "', @ReportDate='" + ReportDt + "'";
                SqlCommand cmd = new SqlCommand(q, con);
                SqlDataAdapter DA = new SqlDataAdapter(cmd);
                DataSet DS = new DataSet();
                DA.Fill(DS, "Core_SOA_Report_Tenant");
                DS.WriteXmlSchema(Server.MapPath("~/Reports/dsReport.xsd"));


                System.IO.StreamReader xmlStream = new System.IO.StreamReader(Server.MapPath("~/Reports/dsReport.xsd"));
                DataSet dataSet = new DataSet();
                dataSet.ReadXmlSchema(xmlStream);
                xmlStream.Close();

                DBClass.SetLog("Entered ExportPDF");
                ReportDocument rd = new ReportDocument();
                rd.Load(Path.Combine(Server.MapPath("~/Reports/rptFiles/"), "SOAReport.rpt"));
                DBClass.SetLog("ExportPDF. Report File path = " + Path.Combine(Server.MapPath("~/Reports/rptFiles/"), "SOAReport.rpt"));

                DBClass.SetLog("@TCNo = " + TcNo);
                DBClass.SetLog("@ReportDate = " + ReportDt);

                rd.SetDataSource(DS);

                string path = Server.MapPath("~/Reports/SOAFiles.pdf");
                FConvert.LogFile("CrystalPrint.log", "Path:Location of PDF" + path);
                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                }
                rd.ExportToDisk(ExportFormatType.PortableDocFormat, path);
                DBClass.SetLog("path = " + path);
                if (path == "")
                {
                    DBClass.SetLog("no path found");
                    System.IO.File.Create(Server.MapPath("~/Reports/SOAFiles.pdf"));
                }
                return Json(path, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                FConvert.LogFile("CrystalPrint.log", "Error in getting report details: " + ex.Message + "\n" + ex.InnerException);
                return null;
            }
        }
        public ActionResult AgeingIndex(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module AgeingIndex with CompanyId = " + CompanyId);
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        public ActionResult OSIndex(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module OSIndex with CompanyId = " + CompanyId);
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        public ActionResult getTenantslist2(int cid, string searchtext, int Property)
        {
            string retrievequery = string.Format($@"exec pCore_CommonSp @Operation=SearchTenants2, @p2='{searchtext}',@p1={Property}");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
            string JSONString = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(JSONString, JsonRequestBehavior.AllowGet);
        }
        public ActionResult getPropertylist(int cid, string searchtext, int TenantId)
        {
            string retrievequery = string.Format($@"exec pCore_CommonSp @Operation=SearchProperty, @p2='{searchtext}',@p1={TenantId}");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
            string JSONString = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(JSONString, JsonRequestBehavior.AllowGet);
        }
        public ActionResult getAccountlist(int cid, string searchtext)
        {
            string retrievequery = string.Format($@"exec pCore_CommonSp @Operation=SearchAccount, @p2='{searchtext}'");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
            string JSONString = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(JSONString, JsonRequestBehavior.AllowGet);
        }
        public ActionResult getPGlist(int cid, string searchtext)
        {
            string retrievequery = string.Format($@"exec pCore_CommonSp @Operation=SearchPropertyGrp, @p2='{searchtext}'");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);
            string JSONString = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(JSONString, JsonRequestBehavior.AllowGet);
        }
        public ActionResult AgeingReport(int CompanyId, string ReportDate, int Tenant, int Property, int Account, string TenantName, string PropertyName, string AccountName)
        {
            try
            {
                DBClass.SetLog("Getting Report View. CompanyId = " + CompanyId.ToString() + " ReportDate = " + ReportDate + " Tenant = " + Tenant + " Property = " + Property + " Account = " + Account);
                ReportData _reportData = new ReportData();
                Search search = new Search();
                _reportData._searchObj = search;
                string retrievequery = string.Format($@"exec Tenant_Ageing_report @Property={Property}, @Tenant={Tenant}, @Account={Account},@ReportDate='{ReportDate}'");
                DBClass.SetLog("AgeingReport retrievequery = " + retrievequery);
                DataSet ds1 = DBClass.GetData(retrievequery, CompanyId, ref errors1);
                if (ds1 != null)
                {
                    DBClass.SetLog("Getting Report View. GetReportData DataSet is not null");
                    if (ds1.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. GetReportData DataSet Table count >0 ");
                        List<AgeingData> listobj = new List<AgeingData>();
                        if (ds1.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow dr1 in ds1.Tables[0].Rows)
                            {
                                listobj.Add(new AgeingData
                                {
                                    DocDate = dr1["DocDate"].ToString(),
                                    DocNo = dr1["sVoucherNo"].ToString(),
                                    Tenant_Name = dr1["TenantName"].ToString(),
                                    Tenant_Code = dr1["TenantCode"].ToString(),
                                    TC_No = dr1["TC_No"].ToString(),
                                    Balance = Convert.ToDecimal(dr1["Balance Amount"].ToString()),
                                    Days30 = Convert.ToDecimal(dr1["0-30 Days"].ToString()),
                                    Days60 = Convert.ToDecimal(dr1["30-60 Days"].ToString()),
                                    Days90 = Convert.ToDecimal(dr1["60-90 Days"].ToString()),
                                    Days120 = Convert.ToDecimal(dr1["90-120 Days"].ToString()),
                                    Days150 = Convert.ToDecimal(dr1["120-150 Days"].ToString()),
                                    Days180 = Convert.ToDecimal(dr1["150-180 Days"].ToString()),
                                    Days360 = Convert.ToDecimal(dr1["180-360 Days"].ToString()),
                                    Days_360 = Convert.ToDecimal(dr1[">360 Days"].ToString()),

                                    Balance2 = dr1["Balance Amount2"].ToString(),
                                    Days302 = dr1["0-30 Days2"].ToString(),
                                    Days602 = dr1["30-60 Days2"].ToString(),
                                    Days902 = dr1["60-90 Days2"].ToString(),
                                    Days1202 = dr1["90-120 Days2"].ToString(),
                                    Days1502 = dr1["120-150 Days2"].ToString(),
                                    Days1802 = dr1["150-180 Days2"].ToString(),
                                    Days3602 = dr1["180-360 Days2"].ToString(),
                                    Days_3602 = dr1[">360 Days2"].ToString(),

                                    PDCNo = dr1["PDCNo"].ToString(),
                                    sChequeNo = dr1["sChequeNo"].ToString(),
                                    pdccount = Convert.ToInt32(dr1["pdccount"].ToString()),
                                });
                            }

                            _reportData._ageingData = listobj;
                            DBClass.SetLog("Getting Report View. GetReportData body data is ready");
                        }
                    }
                }
                search.Tenant = TenantName;
                search.PropertyName = PropertyName;
                search.AccountName = AccountName;
                search.CompanyId = CompanyId;
                search.ReportDate = ReportDate;

                TempData["ReportDate"] = ReportDate;
                TempData["ReportData"] = _reportData;
                TempData["CompanyId"] = CompanyId;
                TempData.Keep();
                return View(_reportData);
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult OSReport2(int CompanyId, string ReportDate, int Tenant, int Property, int Account, string TenantName, string PropertyName, string AccountName)
        {
            try
            {
                DBClass.SetLog("Getting Report View. CompanyId = " + CompanyId.ToString() + " ReportDate = " + ReportDate + " Tenant = " + Tenant + " Property = " + Property + " PropertyGroup = " + Account);
                ReportData _reportData = new ReportData();
                Search search = new Search();
                _reportData._searchObj = search;
                string retrievequery = string.Format($@"exec Tenant_OS_report @Property='{Property}', @Tenant='{Tenant}', @PropertyGrp='{Account}',@ReportDate='{ReportDate}'");
                DBClass.SetLog("OSReport retrievequery = " + retrievequery);
                DataSet ds1 = DBClass.GetData(retrievequery, CompanyId, ref errors1);
                if (ds1 != null)
                {
                    DBClass.SetLog("Getting Report View. GetReportData DataSet is not null");
                    if (ds1.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. GetReportData DataSet Table count >0 ");
                        List<ListData> listobj = new List<ListData>();
                        if (ds1.Tables[0].Rows.Count > 0)
                        {
                            search.ReportDate = ds1.Tables[0].Rows[0]["ReportDate"].ToString();
                            foreach (DataRow dr1 in ds1.Tables[0].Rows)
                            {
                                listobj.Add(new ListData
                                {
                                    Account = dr1["Tenant"].ToString(),
                                    balance = Convert.ToDecimal(dr1["Balance"].ToString()),
                                    Debit = Convert.ToDecimal(dr1["Debit"].ToString()),
                                    Credit = Convert.ToDecimal(dr1["Credit"].ToString()),
                                });
                            }

                            _reportData._listData = listobj;
                            DBClass.SetLog("Getting Report View. GetReportData body data is ready");
                        }
                    }
                }
                else
                {
                    return Json("Success", JsonRequestBehavior.AllowGet);
                }
                search.Tenant = TenantName;
                search.PropertyName = PropertyName;
                search.AccountName = AccountName;
                search.CompanyId = CompanyId;
                Session["OSReportData"] = _reportData;
                if (_reportData == null)
                {
                    return Json("No Data", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json("Success", JsonRequestBehavior.AllowGet);
                }
                DBClass.SetLog("Getting Report View. OSReportData body data is ready");
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult OSReport()
        {
            ReportData _data = (ReportData)Session["OSReportData"];
            return View(_data);
        }
        public FileResult OSExcelGenerate()
        {
            ReportData _data = (ReportData)Session["OSReportData"];
            Search _head = new Search();
            _head = _data._searchObj;
            List<ListData> _list = new List<ListData>();
            _list = _data._listData;
            System.Data.DataTable data = new System.Data.DataTable("Tenant Outstanding Report");
            #region DataColumns
            data.Columns.Add("Tenant", typeof(string));
            data.Columns.Add("Debit", typeof(decimal));
            data.Columns.Add("Credit", typeof(decimal));
            data.Columns.Add("Balance", typeof(decimal));
            #endregion
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Tenant Outstanding Report");
                var dataTable = data;



                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 5));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Tenant Outstanding Report";

                int r = 3;
                int cell = 2;

                ws.Cell(r, 2).Value = "Property Group";
                ws.Cell(r, 3).Value = _head.AccountName;
                ws.Cell(r, 4).Value = "Property";
                ws.Cell(r, 5).Value = _head.PropertyName;

                r = 4;
                ws.Cell(r, 2).Value = "Tenant";
                ws.Cell(r, 3).Value = _head.Tenant;
                ws.Cell(r, 4).Value = "Report Date";
                ws.Cell(r, 5).Value = _head.ReportDate;

                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 5));
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                r = 6;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "Tenant";
                    ws.Cell(r, cell++).Value = "Debit";
                    ws.Cell(r, cell++).Value = "Credit";
                    ws.Cell(r, cell++).Value = "Balance";
                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r, 5));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int c = 2;

                #region TableLoop
                foreach (var obj in _list)
                {
                    c = 2;
                    r++;
                    ws.Cell(r, c++).Value = obj.Account;
                    ws.Cell(r, c++).Value = obj.Debit.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Credit.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.balance.ToString("N", new CultureInfo("en-US"));
                }

                #endregion

                TableRange = ws.Range(ws.Cell(7, 2), ws.Cell(r, 5));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(ws.Cell(7, 2), ws.Cell(r, 2)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(ws.Cell(7, 3), ws.Cell(r, 5)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Range(ws.Cell(7, 3), ws.Cell(r, 5)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "OSReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        [HttpPost]
        public FileResult AgeingExcelGenerate()
        {

            #region TempData
            ReportData _data = new ReportData();
            _data = (ReportData)TempData["ReportData"];
            Search _head = new Search();
            _head = _data._searchObj;
            TempData.Keep();
            #endregion

            System.Data.DataTable data = new System.Data.DataTable("Tenant Ageing Report");
            #region DataColumns
            data.Columns.Add("Date", typeof(string));
            data.Columns.Add("Doc No", typeof(string));
            data.Columns.Add("Tenant Name", typeof(string));
            data.Columns.Add("Tenant Code", typeof(string));
            data.Columns.Add("TC No (Contract No)", typeof(string));
            data.Columns.Add("Balance", typeof(decimal));
            data.Columns.Add("0-30 Days", typeof(decimal));
            data.Columns.Add("30-60 Days", typeof(decimal));
            data.Columns.Add("60-90 Days", typeof(decimal));
            data.Columns.Add("90-120 Days", typeof(decimal));
            data.Columns.Add("120-150 Days", typeof(decimal));
            data.Columns.Add("150-180Days", typeof(decimal));
            data.Columns.Add("180-360 Days", typeof(decimal));
            data.Columns.Add(">360 Days", typeof(decimal));
            #endregion


            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Tenant Ageing Report");
                var dataTable = data;



                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 15));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Tenant Ageing Report";

                int r = 3;
                ws.Cell(r, 2).Value = "Property";
                ws.Range(ws.Cell(r, 3), ws.Cell(r, 6)).Merge().Value = _head.PropertyName;
                ws.Cell(r, 9).Value = "Account";
                ws.Range(ws.Cell(r, 10), ws.Cell(r, 15)).Merge().Value = _head.AccountName;

                r = 4;
                int cell = 2;
                ws.Cell(r, 2).Value = "Tenant";
                ws.Range(ws.Cell(r, 3), ws.Cell(r, 6)).Merge().Value = _head.Tenant;
                ws.Cell(r, 9).Value = "Date";
                ws.Range(ws.Cell(r, 10), ws.Cell(r, 15)).Merge().Value = _head.ReportDate;

                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 15));
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                r = 7;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "Date";
                    ws.Cell(r, cell++).Value = "Doc No";
                    ws.Cell(r, cell++).Value = "Tenant Name";
                    ws.Cell(r, cell++).Value = "Tenant Code";
                    ws.Cell(r, cell++).Value = "TC No (Contract No)";
                    ws.Cell(r, cell++).Value = "Balance";
                    ws.Cell(r, cell++).Value = "0-30 Days";
                    ws.Cell(r, cell++).Value = "30-60 Days";
                    ws.Cell(r, cell++).Value = "60-90 Days";
                    ws.Cell(r, cell++).Value = "90-120 Days";
                    ws.Cell(r, cell++).Value = "120-150 Days";
                    ws.Cell(r, cell++).Value = "150-180Days";
                    ws.Cell(r, cell++).Value = "180-360 Days";
                    ws.Cell(r, cell++).Value = ">360 Days";

                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r, 15));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                int c = 2;

                #region TableLoop
                decimal Days30 = 0;
                decimal Days60 = 0;
                decimal Days90 = 0;
                decimal Days120 = 0;
                decimal Days150 = 0;
                decimal Days180 = 0;
                decimal Days360 = 0;
                decimal Days_360 = 0;
                decimal bal = 0;
                int count = 1;
                List<AgeingData> _list = new List<AgeingData>();
                _list = _data._ageingData;
                foreach (var obj in _list)
                {
                    Days30 = Convert.ToDecimal(obj.Days30) + Days30;
                    bal = Convert.ToDecimal(obj.Balance) + bal;
                    Days60 = Convert.ToDecimal(obj.Days60) + Days60;
                    Days90 = Convert.ToDecimal(obj.Days90) + Days90;
                    Days120 = Convert.ToDecimal(obj.Days120) + Days120;
                    Days150 = Convert.ToDecimal(obj.Days150) + Days150;
                    Days180 = Convert.ToDecimal(obj.Days180) + Days180;
                    Days360 = Convert.ToDecimal(obj.Days360) + Days360;
                    Days_360 = Convert.ToDecimal(obj.Days_360) + Days_360;
                    c = 2;
                    r++;
                    ws.Cell(r, c++).Value = obj.DocDate;
                    ws.Cell(r, c++).Value = obj.DocNo;
                    ws.Cell(r, c++).Value = obj.Tenant_Name;
                    ws.Cell(r, c++).Value = obj.Tenant_Code;
                    ws.Cell(r, c++).Value = obj.TC_No;
                    ws.Cell(r, c++).Value = obj.Balance2;
                    ws.Cell(r, c++).Value = obj.Days302;
                    ws.Cell(r, c++).Value = obj.Days602;
                    ws.Cell(r, c++).Value = obj.Days902;
                    ws.Cell(r, c++).Value = obj.Days1202;
                    ws.Cell(r, c++).Value = obj.Days1502;
                    ws.Cell(r, c++).Value = obj.Days1802;
                    ws.Cell(r, c++).Value = obj.Days3602;
                    ws.Cell(r, c++).Value = obj.Days_3602;
                }


                //Grand Total Row
                r++;
                c = 2;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 5)).Merge().Value = "Total";
                ws.Cell(r, 6).Value = bal.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 7).Value = Days30.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 8).Value = Days60.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 9).Value = Days90.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 10).Value = Days120.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 11).Value = Days150.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 12).Value = Days180.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 13).Value = Days360.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, 14).Value = Days_360.ToString("N", new CultureInfo("en-US"));

                ws.Range("B" + r + ":Z" + r + "").Style.Font.Bold = true;
                r++;
                int pdcrow = r;
                if (_data._ageingData[0].pdccount > 0)
                {
                    ws.Range(ws.Cell(r, 2), ws.Cell(r, 15)).Merge();
                    r = r + 1;
                    wsReportNameHeaderRange = ws.Range(ws.Cell(r++, 2), ws.Cell(r++, 15));
                    wsReportNameHeaderRange.Style.Font.Bold = true;
                    wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.White;
                    wsReportNameHeaderRange.Merge();
                    wsReportNameHeaderRange.Value = "Summary";

                    wsReportNameHeaderRange = ws.Range(ws.Cell(r++, 2), ws.Cell(r++, 15));
                    wsReportNameHeaderRange.Style.Font.Bold = true;
                    wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.White;
                    wsReportNameHeaderRange.Merge();
                    wsReportNameHeaderRange.Value = "Uncleared PDC";
                    cell = 2;
                    ws.Range(ws.Cell(r, 2), ws.Cell(r, 15)).Merge();
                    r = r + 1;
                    #region Headers
                    ws.Range(ws.Cell(r, cell), ws.Cell(r, 4)).Merge().Value = "Doc No";
                    ws.Range(ws.Cell(r, 5), ws.Cell(r, 7)).Merge().Value = "Cheque No";
                    ws.Range(ws.Cell(r, 8), ws.Cell(r, 11)).Merge().Value = "Cheque Date";
                    ws.Range(ws.Cell(r, 12), ws.Cell(r, 15)).Merge().Value = "Amount";
                    #endregion
                    foreach (var obj in _list)
                    {
                        if (obj.PDCNo != "0")
                        {
                            c = 2;
                            r++;
                            ws.Range(ws.Cell(r, c), ws.Cell(r, 4)).Merge().Value = obj.DocNo;
                            ws.Range(ws.Cell(r, 5), ws.Cell(r, 7)).Merge().Value = obj.sChequeNo;
                            ws.Range(ws.Cell(r, 8), ws.Cell(r, 11)).Merge().Value = obj.DocDate;
                            ws.Range(ws.Cell(r, 12), ws.Cell(r, 15)).Merge().Value = obj.Balance.ToString("N", new CultureInfo("en-US"));
                        }
                    }

                }

                #endregion

                TableRange = ws.Range(ws.Cell(8, 2), ws.Cell(r, 15));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;



                ws.Range(ws.Cell(8, 7), ws.Cell(pdcrow - 1, 15)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TenantAgeingReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        public DataSet GetOccupancyAllFilters(int cid)
        {
            DataSet ds = new DataSet();
            try
            {
                string retrievequery = string.Format(@"exec pCore_CommonSp @Operation=getPropertyGrp exec pCore_CommonSp @Operation=getProperty exec pCore_CommonSp @Operation=getUnitMaster exec pCore_CommonSp @Operation=getUnitType exec pCore_CommonSp @Operation=getUsage exec pCore_CommonSp @Operation=getTC");
                DBClass.SetLog("retrievequery = " + retrievequery);
                ds = DBClass.GetData(retrievequery, cid, ref errors1);
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
            }
            return ds;
        }
        public List<SelectOptions> getPG(DataTable dt)
        {
            List<SelectOptions> containers = new List<SelectOptions>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                containers.Add(new SelectOptions()
                {
                    Name = dt.Rows[i]["sName"].ToString(),
                    Id = dt.Rows[i]["iMasterId"].ToString(),
                    //Code = dt.Rows[i]["sCode"].ToString(),
                });
            }
            return containers;
        }
        public List<SelectOptions> getProperty(DataTable dt)
        {
            List<SelectOptions> containers = new List<SelectOptions>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                containers.Add(new SelectOptions()
                {
                    Name = dt.Rows[i]["sName"].ToString(),
                    Id = dt.Rows[i]["iMasterId"].ToString(),
                    FId = dt.Rows[i]["PropertyGroup"].ToString(),
                });
            }
            return containers;
        }
        public List<SelectOptions> getUnit(DataTable dt)
        {
            List<SelectOptions> containers = new List<SelectOptions>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                containers.Add(new SelectOptions()
                {
                    Name = dt.Rows[i]["sName"].ToString(),
                    Id = dt.Rows[i]["iMasterId"].ToString(),
                    FId = dt.Rows[i]["Property"].ToString(),
                    Code = dt.Rows[i]["UsageType"].ToString(),
                    Extra = dt.Rows[i]["UnitType"].ToString(),
                });
            }
            return containers;
        }
        public List<SelectOptions> getUType(DataTable dt)
        {
            List<SelectOptions> containers = new List<SelectOptions>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                containers.Add(new SelectOptions()
                {
                    Name = dt.Rows[i]["sName"].ToString(),
                    Id = dt.Rows[i]["iMasterId"].ToString(),
                });
            }
            return containers;
        }
        public List<SelectOptions> getUsage(DataTable dt)
        {
            List<SelectOptions> containers = new List<SelectOptions>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                containers.Add(new SelectOptions()
                {
                    Name = dt.Rows[i]["sName"].ToString(),
                    Id = dt.Rows[i]["iMasterId"].ToString(),
                });
            }
            return containers;
        }
        public List<SelectOptions> getTC(DataTable dt)
        {
            List<SelectOptions> containers = new List<SelectOptions>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                containers.Add(new SelectOptions()
                {
                    Name = dt.Rows[i]["sName"].ToString(),
                    Id = dt.Rows[i]["iMasterId"].ToString(),
                    FId = dt.Rows[i]["UnitMaster"].ToString(),
                });
            }
            return containers;
        }
        public ActionResult OccupancyIndex(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module with CompanyId = " + CompanyId);
                DataSet ds = GetOccupancyAllFilters(CompanyId);
                var pg = getPG(ds.Tables[0]);
                var pro = getProperty(ds.Tables[1]);
                var unit = getUnit(ds.Tables[2]);
                var UType = getUType(ds.Tables[3]);
                var Usage = getUsage(ds.Tables[4]);
                var tc = getTC(ds.Tables[5]);
                ViewBag.propertygrp = pg;
                ViewBag.property = pro;
                ViewBag.unit = unit;
                ViewBag.unittype = UType;
                ViewBag.usage = Usage;
                ViewBag.TCno = tc;
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        [HttpPost]
        public ActionResult OccupanyReport2(OccupancyFilter obj)
        {
            try
            {
                OccupancyCls _cls = new OccupancyCls();
                OccupancyFilter _filter = new OccupancyFilter();
                _cls._filter = obj;
                int CompanyId = Convert.ToInt32(obj.CompanyId);
                string retrievequery = string.Format($@"exec OccupanyReport @PropertyGrps='{obj.PropertyGrp}', @Proeprtys='{obj.Property}', @UnitIds='{obj.Unit}',@TCNo='{obj.TCNoid}',@UnitType='{obj.UnitType}',@Usage='{obj.Usage}',@Sqft={obj.Sqft},@SqftTo={obj.SqftTo},@SDate='{obj.FromDate}',@EDate='{obj.ToDate}'");
                DBClass.SetLog("OccupanyReport retrievequery = " + retrievequery);
                DataSet ds1 = DBClass.GetData(retrievequery, CompanyId, ref errors1);
                if (ds1 != null)
                {
                    DBClass.SetLog("Getting Report View. OccupanyReport DataSet is not null");
                    if (ds1.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. OccupanyReport DataSet Table count >0 ");
                        List<OccupancyList> listobj = new List<OccupancyList>();
                        if (ds1.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow dr1 in ds1.Tables[0].Rows)
                            {
                                listobj.Add(new OccupancyList
                                {
                                    PropertyGrp = dr1["PGname"].ToString(),
                                    Property = dr1["PName"].ToString(),
                                    Unit = dr1["UnitName"].ToString(),
                                    TCNo = dr1["TCName"].ToString(),
                                    FromDate = dr1["Fromdate"].ToString(),
                                    ToDate = dr1["Enddate"].ToString(),
                                    UnitType = dr1["UnitTypeName"].ToString(),
                                    Usage = dr1["UsageName"].ToString(),
                                    Sqft = Convert.ToDecimal(dr1["Area"].ToString()),
                                    AccPrdDays = Convert.ToInt32(dr1["AccPrdDays"].ToString()),
                                    ContractAmt = Convert.ToDecimal(dr1["ContractValue"].ToString()),
                                    AmFrom = dr1["Am_start"].ToString(),
                                    AmTo = dr1["Am_End"].ToString(),
                                    AmDays = Convert.ToInt32(dr1["AmDays"].ToString()),
                                    AmAmt = Convert.ToDecimal(dr1["AmAmt"].ToString()),
                                    Status = dr1["Leasestatus"].ToString(),
                                    dayRent = Convert.ToDecimal(dr1["rentperday"].ToString()),
                                    sqRent = Convert.ToDecimal(dr1["rentpersqft"].ToString()),
                                    vacdays = Convert.ToInt32(dr1["vaccantdays"].ToString()),
                                    AnnualRent = Convert.ToDecimal(dr1["AnnualRent"].ToString()),
                                    vacLoss = Convert.ToDecimal(dr1["rentloss"].ToString()),
                                    UnAm_Amt = Convert.ToDecimal(dr1["UnAm_rent"].ToString()),
                                });
                            }

                            _cls._list = listobj;
                            Session["OccupanyData"] = _cls;
                            if (_cls == null)
                            {
                                return Json("No Data", JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                return Json("Success", JsonRequestBehavior.AllowGet);
                            }
                            DBClass.SetLog("Getting Report View. OccupanyReportData body data is ready");
                        }
                        else
                        {
                            return Json("No Data", JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        return Json("No Data", JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    return Json("No Data", JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult OccupanyReport()
        {
            OccupancyCls _data = (OccupancyCls)Session["OccupanyData"];
            return View(_data);
        }
        [HttpPost]
        public FileResult OccupancyExcel()
        {
            OccupancyCls _data = (OccupancyCls)Session["OccupanyData"]; 
            OccupancyFilter _head = new OccupancyFilter();
            _head = _data._filter;
            List<OccupancyList> _list = new List<OccupancyList>();
            _list = _data._list;
            System.Data.DataTable data = new System.Data.DataTable("Occupancy Report");
            #region DataColumns
            data.Columns.Add("Property Group", typeof(string));
            data.Columns.Add("Property Name", typeof(string));
            data.Columns.Add("TC No", typeof(string));
            data.Columns.Add("Unit No", typeof(string));
            data.Columns.Add("Unit Type", typeof(string));
            data.Columns.Add("Unit Residential/ Commercial", typeof(string));
            data.Columns.Add("Unit Area Sq Ft", typeof(decimal));
            data.Columns.Add("Accounting Period Run From", typeof(string));
            data.Columns.Add("Accounting Period Run To", typeof(string));
            data.Columns.Add("Accounting Period Days", typeof(int));
            data.Columns.Add("Contract Value (tc no value)", typeof(decimal));
            data.Columns.Add("Amortisation From( tc no start dt)", typeof(string));
            data.Columns.Add("Amortisation To(tc termination dt)", typeof(string));
            data.Columns.Add("Amortisation Days", typeof(int));
            data.Columns.Add("Amortised Rent Amount", typeof(decimal));
            data.Columns.Add("Status (Leased / Not-leased) for the period", typeof(string));
            data.Columns.Add("Rent Per Day", typeof(decimal));
            data.Columns.Add("Rent Per Sq Ft Per Day", typeof(decimal));
            data.Columns.Add("(Take only if TC no is blank) Vacant Period", typeof(decimal));
            data.Columns.Add("Annual Rent (Unit Master)", typeof(decimal));
            data.Columns.Add("Estimated Vacancy Period Rent Loss", typeof(decimal));
            data.Columns.Add("Unamortised Rent Amount as on accounting period end", typeof(decimal));

            #endregion


            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Occupancy Report");
                var dataTable = data;



                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 23));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Occupancy Report";

                int r = 3;
                int cell = 2;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 6)).Merge().Value = "Property Group";
                ws.Range(ws.Cell(r, 7), ws.Cell(r, 11)).Merge().Value = _head.PropertyGrpName;
                ws.Range(ws.Cell(r, 14), ws.Cell(r, 18)).Merge().Value = "Property";
                ws.Range(ws.Cell(r, 19), ws.Cell(r, 23)).Merge().Value = _head.PropertyName;

                r = 4;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 6)).Merge().Value = "Unit Type";
                ws.Range(ws.Cell(r, 7), ws.Cell(r, 11)).Merge().Value = _head.UnitTypeName;
                ws.Range(ws.Cell(r, 14), ws.Cell(r, 18)).Merge().Value = "Usage";
                ws.Range(ws.Cell(r, 19), ws.Cell(r, 23)).Merge().Value = _head.UsageName;

                r = 5;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 6)).Merge().Value = "Unit";
                ws.Range(ws.Cell(r, 7), ws.Cell(r, 11)).Merge().Value = _head.UnitName;
                ws.Range(ws.Cell(r, 14), ws.Cell(r, 18)).Merge().Value = "TC No";
                ws.Range(ws.Cell(r, 19), ws.Cell(r, 23)).Merge().Value = _head.TcNos;

                r = 6;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 6)).Merge().Value = "From";
                ws.Range(ws.Cell(r, 7), ws.Cell(r, 11)).Merge().Value = _list[0].FromDate;
                ws.Range(ws.Cell(r, 14), ws.Cell(r, 18)).Merge().Value = "To";
                ws.Range(ws.Cell(r, 19), ws.Cell(r, 23)).Merge().Value = _list[0].ToDate;

                r = 7;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 6)).Merge().Value = "Sq. Ft From";
                ws.Range(ws.Cell(r, 7), ws.Cell(r, 11)).Merge().Value = _head.Sqft;
                ws.Range(ws.Cell(r, 14), ws.Cell(r, 18)).Merge().Value = "Sq.Ft To";
                ws.Range(ws.Cell(r, 19), ws.Cell(r, 23)).Merge().Value = _head.SqftTo;


                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 23));
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(ws.Cell(r, 7), ws.Cell(r, 11)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(r, 19), ws.Cell(r, 23)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(3, 2), ws.Cell(7, 6)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Range(ws.Cell(3, 14), ws.Cell(7, 18)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                ws.Range(ws.Cell(3, 7), ws.Cell(7, 11)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(ws.Cell(3, 19), ws.Cell(7, 23)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                r = 9;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "Property Group";
                    ws.Cell(r, cell++).Value = "Property Name";
                    ws.Cell(r, cell++).Value = "TC No";
                    ws.Cell(r, cell++).Value = "Unit No";
                    ws.Cell(r, cell++).Value = "Unit Type";
                    ws.Cell(r, cell++).Value = "Unit Residential/ Commercial";
                    ws.Cell(r, cell++).Value = "Unit Area Sq Ft";
                    ws.Cell(r, cell++).Value = "Accounting Period Run From";
                    ws.Cell(r, cell++).Value = "Accounting Period Run To";
                    ws.Cell(r, cell++).Value = "Accounting Period Days";
                    ws.Cell(r, cell++).Value = "Contract Value (tc no value)";
                    ws.Cell(r, cell++).Value = "Amortisation From( tc no start dt)";
                    ws.Cell(r, cell++).Value = "Amortisation To(tc termination dt)";
                    ws.Cell(r, cell++).Value = "Amortisation Days";
                    ws.Cell(r, cell++).Value = "Amortised Rent Amount";
                    ws.Cell(r, cell++).Value = "Status (Leased / Not-leased) for the period";
                    ws.Cell(r, cell++).Value = "Rent Per Day";
                    ws.Cell(r, cell++).Value = "Rent Per Sq Ft Per Day";
                    ws.Cell(r, cell++).Value = "(Take only if TC no is blank) Vacant Period";
                    ws.Cell(r, cell++).Value = "Annual Rent (Unit Master)";
                    ws.Cell(r, cell++).Value = "Estimated Vacancy Period Rent Loss";
                    ws.Cell(r, cell++).Value = "Unamortised Rent Amount as on accounting period end";
                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r,23));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int c = 2;

                #region TableLoop
                foreach (var obj in _list)
                {
                    c = 2;
                    r++;
                    ws.Cell(r, c++).Value = obj.PropertyGrp;
                    ws.Cell(r, c++).Value = obj.Property;
                    ws.Cell(r, c++).Value = obj.TCNo;
                    ws.Cell(r, c++).Value = obj.Unit;
                    ws.Cell(r, c++).Value = obj.UnitType;
                    ws.Cell(r, c++).Value = obj.Usage;
                    ws.Cell(r, c++).Value = obj.Sqft.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.FromDate;
                    ws.Cell(r, c++).Value = obj.ToDate;
                    ws.Cell(r, c++).Value = obj.AccPrdDays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.ContractAmt.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.AmFrom;
                    ws.Cell(r, c++).Value = obj.AmTo;
                    ws.Cell(r, c++).Value = obj.AmDays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.AmAmt.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Status;
                    ws.Cell(r, c++).Value = obj.dayRent.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.sqRent.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.vacdays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.AnnualRent.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.vacLoss.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.UnAm_Amt.ToString("N", new CultureInfo("en-US"));
                }

                #endregion

                TableRange = ws.Range(ws.Cell(10, 2), ws.Cell(r, 23));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range(ws.Cell(10, 8), ws.Cell(r, 8)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(10, 11), ws.Cell(r, 11)).Style.NumberFormat.Format = "0";
                ws.Range(ws.Cell(10, 12), ws.Cell(r, 12)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(10, 15), ws.Cell(r, 15)).Style.NumberFormat.Format = "0";
                ws.Range(ws.Cell(10, 16), ws.Cell(r, 16)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(10, 18), ws.Cell(r, 23)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "OccupancyReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        public ActionResult ExpenseIndex(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module with CompanyId = " + CompanyId);
                DataSet ds = GetOccupancyAllFilters(CompanyId);
                var pg = getPG(ds.Tables[0]);
                var pro = getProperty(ds.Tables[1]);
                var unit = getUnit(ds.Tables[2]);
                var UType = getUType(ds.Tables[3]);
                ViewBag.propertygrp = pg;
                ViewBag.property = pro;
                ViewBag.unit = unit;
                ViewBag.unittype = UType;
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        public ActionResult ExpenseReport2(ExpFilter obj)
        {
            try
            {
                ExpenseCls _cls = new ExpenseCls();
                ExpFilter _filter = new ExpFilter();
                _cls._filter = obj;
                int CompanyId = Convert.ToInt32(obj.CompanyId);
                string retrievequery = string.Format($@"exec ExpenseReport @PropertyGrps='{obj.PropertyGrp}', @Proeprtys='{obj.Property}', @UnitIds='{obj.Unit}',@UnitType='{obj.UnitType}',@SDate='{obj.FromDate}',@EDate='{obj.ToDate}'");
                DBClass.SetLog("ExpenseReport retrievequery = " + retrievequery);
                DataSet ds1 = DBClass.GetData(retrievequery, CompanyId, ref errors1);
                if (ds1 != null)
                {
                    DBClass.SetLog("Getting Report View. ExpenseReport DataSet is not null");
                    if (ds1.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. ExpenseReport DataSet Table count >0 ");
                        List<ExpList> listobj = new List<ExpList>();
                        if (ds1.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow dr1 in ds1.Tables[0].Rows)
                            {
                                listobj.Add(new ExpList
                                {
                                    PropertyGrp = dr1["PropertyGrpName"].ToString(),
                                    Property = dr1["PropertyName"].ToString(),
                                    Unit = dr1["Unitno"].ToString(),
                                    FromDate = dr1["Fromdate"].ToString(),
                                    ToDate = dr1["Enddate"].ToString(),
                                    UnitType = dr1["UnitType"].ToString(),
                                    Usage = dr1["Usage"].ToString(),
                                    Sqft = dr1["Area"].ToString(),
                                    AccPrdDays = Convert.ToInt32(dr1["AccPrdDays"].ToString()),
                                    Account = dr1["expHead"].ToString(),
                                    Amt = Convert.ToDecimal(dr1["Amt"].ToString()),
                                });
                            }

                            _cls._list = listobj;
                            Session["ExpData"] = _cls;
                            if (_cls == null)
                            {
                                return Json("No Data", JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                return Json("Success", JsonRequestBehavior.AllowGet);
                            }
                            DBClass.SetLog("Getting Report View. ExpenseReportData body data is ready");
                        }
                        else
                        {
                            return Json("No Data", JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        return Json("No Data", JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    return Json(errors1, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult ExpenseReport()
        {
            ExpenseCls _data = (ExpenseCls)Session["ExpData"];
            return View(_data);
        }
        public FileResult ExpenseExcel()
        {
            ExpenseCls _data = (ExpenseCls)Session["ExpData"];
            ExpFilter _head = new ExpFilter();
            _head = _data._filter;
            List<ExpList> _list = new List<ExpList>();
            _list = _data._list;
            System.Data.DataTable data = new System.Data.DataTable("Expense Report");
            #region DataColumns
            data.Columns.Add("Property Group", typeof(string));
            data.Columns.Add("Property Name", typeof(string));
            data.Columns.Add("Unit No", typeof(string));
            data.Columns.Add("Unit Type", typeof(string));
            data.Columns.Add("Unit Residential/ Commercial", typeof(string));
            data.Columns.Add("Unit Area Sq Ft", typeof(decimal));
            data.Columns.Add("Accounting Period Run From", typeof(string));
            data.Columns.Add("Accounting Period Run To", typeof(string));
            data.Columns.Add("Accounting Period Days", typeof(int));
            data.Columns.Add("Expense Head", typeof(string));
            data.Columns.Add("Expense Amount", typeof(decimal));
            #endregion
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Expense Report");
                var dataTable = data;



                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 12));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Expense Report";

                int r = 3;
                int cell = 2; 

                ws.Cell(r, 2).Value = "Property Group";
                ws.Range(ws.Cell(r, 3), ws.Cell(r, 6)).Merge().Value = _head.PropertyGrpName;
                ws.Cell(r, 8).Value = "Property";
                ws.Range(ws.Cell(r, 9), ws.Cell(r, 12)).Merge().Value = _head.PropertyName;

                r = 4;
                ws.Cell(r, 2).Value = "Unit";
                ws.Range(ws.Cell(r, 3), ws.Cell(r, 6)).Merge().Value = _head.UnitName;
                ws.Cell(r, 8).Value = "Unit Type";
                ws.Range(ws.Cell(r, 9), ws.Cell(r, 12)).Merge().Value = _head.UnitTypeName;

                r = 5;
                ws.Cell(r, 2).Value = "From";
                ws.Range(ws.Cell(r, 3), ws.Cell(r, 6)).Merge().Value = _list[0].FromDate;
                ws.Cell(r, 8).Value = "To";
                ws.Range(ws.Cell(r, 9), ws.Cell(r, 12)).Merge().Value = _list[0].ToDate;

                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 12));
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                r = 7;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "Property Group";
                    ws.Cell(r, cell++).Value = "Property Name";
                    ws.Cell(r, cell++).Value = "Unit No";
                    ws.Cell(r, cell++).Value = "Unit Type";
                    ws.Cell(r, cell++).Value = "Unit Residential/ Commercial";
                    ws.Cell(r, cell++).Value = "Unit Area Sq Ft";
                    ws.Cell(r, cell++).Value = "Accounting Period Run From";
                    ws.Cell(r, cell++).Value = "Accounting Period Run To";
                    ws.Cell(r, cell++).Value = "Accounting Period Days";
                    ws.Cell(r, cell++).Value = "Expense Head";
                    ws.Cell(r, cell++).Value = "Expense Amount";
                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r, 12));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int c = 2;

                #region TableLoop
                foreach (var obj in _list)
                {
                    c = 2;
                    r++;
                    ws.Cell(r, c++).Value = obj.PropertyGrp;
                    ws.Cell(r, c++).Value = obj.Property;
                    ws.Cell(r, c++).Value = obj.Unit;
                    ws.Cell(r, c++).Value = obj.UnitType;
                    ws.Cell(r, c++).Value = obj.Usage;
                    ws.Cell(r, c++).Value = obj.Sqft;
                    ws.Cell(r, c++).Value = obj.FromDate;
                    ws.Cell(r, c++).Value = obj.ToDate;
                    ws.Cell(r, c++).Value = obj.AccPrdDays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Account;
                    ws.Cell(r, c++).Value = obj.Amt.ToString("N", new CultureInfo("en-US"));
                }

                #endregion

                TableRange = ws.Range(ws.Cell(7, 2), ws.Cell(r, 12));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range(ws.Cell(8, 10), ws.Cell(r, 10)).Style.NumberFormat.Format = "0";
                ws.Range(ws.Cell(8, 12), ws.Cell(r, 12)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExpenseReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        public ActionResult LostRevDetIndex(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module with CompanyId = " + CompanyId);
                DataSet ds = GetOccupancyAllFilters(CompanyId);
                var pg = getPG(ds.Tables[0]);
                var pro = getProperty(ds.Tables[1]);
                var unit = getUnit(ds.Tables[2]);
                ViewBag.propertygrp = pg;
                ViewBag.property = pro;
                ViewBag.unit = unit;
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        public ActionResult LostRevDetReport2(LostRevFilter obj)
        {
            try
            {
                LostRevDetCls _cls = new LostRevDetCls();
                LostRevFilter _filter = new LostRevFilter();
                _cls._filter = obj;
                int CompanyId = Convert.ToInt32(obj.CompanyId);
                string retrievequery = string.Format($@"exec LostRevenueDetailReport @PropertyGrps='{obj.PropertyGrp}', @Proeprtys='{obj.Property}', @UnitIds='{obj.Unit}',@SDate='{obj.FromDate}',@EDate='{obj.ToDate}'");
                DBClass.SetLog("ExpenseReport retrievequery = " + retrievequery);
                DataSet ds1 = DBClass.GetData(retrievequery, CompanyId, ref errors1);
                if (ds1 != null)
                {
                    DBClass.SetLog("Getting Report View. ExpenseReport DataSet is not null");
                    if (ds1.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. ExpenseReport DataSet Table count >0 ");
                        List<LostRevDetList> listobj = new List<LostRevDetList>();
                        if (ds1.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow dr1 in ds1.Tables[0].Rows)
                            {
                                listobj.Add(new LostRevDetList
                                {
                                    Property = dr1["PName"].ToString(),
                                    Unit = dr1["UnitName"].ToString(),
                                    AnnualRent = Convert.ToDecimal(dr1["AnnualRent"].ToString()),
                                    LostStartDt = dr1["Am_start"].ToString(),
                                    LostEndDt = dr1["Am_End"].ToString(),
                                    FromDate = dr1["Fromdate"].ToString(),
                                    ToDate = dr1["Enddate"].ToString(),
                                    TotLostRev = Convert.ToDecimal(dr1["rentloss"].ToString()),
                                    Jan = Convert.ToDecimal(dr1["Jan"].ToString()),
                                    TotLostDays = Convert.ToInt32(dr1["vaccantdays"].ToString()),
                                    Feb = Convert.ToDecimal(dr1["Feb"].ToString()),
                                    Mar = Convert.ToDecimal(dr1["Mar"].ToString()),
                                    Apr= Convert.ToDecimal(dr1["Apr"].ToString()),
                                    May= Convert.ToDecimal(dr1["May"].ToString()),
                                    Jun = Convert.ToDecimal(dr1["Jun"].ToString()),
                                    Jul= Convert.ToDecimal(dr1["Jul"].ToString()),
                                    Aug= Convert.ToDecimal(dr1["Aug"].ToString()),
                                    Sep= Convert.ToDecimal(dr1["Sep"].ToString()),
                                    Oct= Convert.ToDecimal(dr1["Oct"].ToString()),
                                    Nov= Convert.ToDecimal(dr1["Nov"].ToString()),
                                    Dec= Convert.ToDecimal(dr1["Dec"].ToString())
                                });
                            }

                            _cls._list = listobj;
                            Session["LostRevDetData"] = _cls;
                            if (_cls == null)
                            {
                                return Json("No Data", JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                return Json("Success", JsonRequestBehavior.AllowGet);
                            }
                            DBClass.SetLog("Getting Report View. ExpenseReportData body data is ready");
                        }
                        else
                        {
                            return Json("No Data", JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        return Json("No Data", JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    return Json(errors1, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult LostRevDetReport()
        {
            LostRevDetCls _data = (LostRevDetCls)Session["LostRevDetData"];
            return View(_data);
        }
        public FileResult LostRevDetExcel()
        {
            LostRevDetCls _data = (LostRevDetCls)Session["LostRevDetData"];
            LostRevFilter _head = new LostRevFilter();
            _head = _data._filter;
            List<LostRevDetList> _list = new List<LostRevDetList>();
            _list = _data._list;
            System.Data.DataTable data = new System.Data.DataTable("Lost Revenue Detail Report");
            #region DataColumns
            data.Columns.Add("Property Name", typeof(string));
            data.Columns.Add("Unit No", typeof(string));
            data.Columns.Add("Annual Rent", typeof(decimal));
            data.Columns.Add("Lost Start Date", typeof(string));
            data.Columns.Add("Lost End Date", typeof(string));
            data.Columns.Add("Total Lost Days", typeof(int));
            data.Columns.Add("Total Revenue Lost", typeof(decimal));
            data.Columns.Add("Jan", typeof(decimal));
            data.Columns.Add("Feb", typeof(decimal));
            data.Columns.Add("Mar", typeof(decimal));
            data.Columns.Add("Apr", typeof(decimal));
            data.Columns.Add("May", typeof(decimal));
            data.Columns.Add("Jun", typeof(decimal));
            data.Columns.Add("Jul", typeof(decimal));
            data.Columns.Add("Aug", typeof(decimal));
            data.Columns.Add("Sep", typeof(decimal));
            data.Columns.Add("Oct", typeof(decimal));
            data.Columns.Add("Nov", typeof(decimal));
            data.Columns.Add("Dec", typeof(decimal));
            #endregion

            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Lost Revenue Detail Report");
                var dataTable = data;
                
                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 20));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Lost Revenue Detail Report";

                int r = 3;
                int cell = 2;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 5)).Merge().Value = "From";
                ws.Range(ws.Cell(r, 6), ws.Cell(r, 10)).Merge().Value = _list[0].FromDate;
                ws.Range(ws.Cell(r, 13), ws.Cell(r, 16)).Merge().Value = "To";
                ws.Range(ws.Cell(r, 17), ws.Cell(r, 20)).Merge().Value = _list[0].ToDate;

                r = 4;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 5)).Merge().Value = "Property Group";
                ws.Range(ws.Cell(r, 6), ws.Cell(r, 10)).Merge().Value = _head.PropertyGrpName;
                ws.Range(ws.Cell(r, 13), ws.Cell(r, 16)).Merge().Value = "Property";
                ws.Range(ws.Cell(r, 17), ws.Cell(r, 20)).Merge().Value = _head.PropertyName;

                r = 5;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 5)).Merge().Value = "Unit";
                ws.Range(ws.Cell(r, 6), ws.Cell(r, 10)).Merge().Value = _head.UnitName;
                ws.Range(ws.Cell(r, 13), ws.Cell(r, 16)).Merge().Value = "";
                ws.Range(ws.Cell(r, 17), ws.Cell(r, 20)).Merge().Value = "";

                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 20));
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(ws.Cell(3, 2), ws.Cell(5, 5)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Range(ws.Cell(3, 13), ws.Cell(5, 16)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                ws.Range(ws.Cell(3, 6), ws.Cell(5, 10)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(ws.Cell(3, 17), ws.Cell(5, 20)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                r = 7;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "Property Name";
                    ws.Cell(r, cell++).Value = "Unit No";
                    ws.Cell(r, cell++).Value = "Annual Rent";
                    ws.Cell(r, cell++).Value = "Lost Start Date";
                    ws.Cell(r, cell++).Value = "Lost End Date";
                    ws.Cell(r, cell++).Value = "Total Lost Days";
                    ws.Cell(r, cell++).Value = "Total Revenue Lost";
                    ws.Cell(r, cell++).Value = "Jan";
                    ws.Cell(r, cell++).Value = "Feb";
                    ws.Cell(r, cell++).Value = "Mar";
                    ws.Cell(r, cell++).Value = "Apr";
                    ws.Cell(r, cell++).Value = "May";
                    ws.Cell(r, cell++).Value = "Jun";
                    ws.Cell(r, cell++).Value = "Jul";
                    ws.Cell(r, cell++).Value = "Aug";
                    ws.Cell(r, cell++).Value = "Sep";
                    ws.Cell(r, cell++).Value = "Oct";
                    ws.Cell(r, cell++).Value = "Nov";
                    ws.Cell(r, cell++).Value = "Dec";
                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r, 20));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int c = 2;

                #region TableLoop
                foreach (var obj in _list)
                {
                    c = 2;
                    r++;
                    ws.Cell(r, c++).Value = obj.Property;
                    ws.Cell(r, c++).Value = obj.Unit;
                    ws.Cell(r, c++).Value = obj.AnnualRent.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.LostStartDt;
                    ws.Cell(r, c++).Value = obj.LostEndDt;
                    ws.Cell(r, c++).Value = obj.TotLostDays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.TotLostRev.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Jan.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Feb.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Mar.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Apr.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.May.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Jun.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Jul.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Aug.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Sep.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Oct.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Nov.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Dec.ToString("N", new CultureInfo("en-US"));
                }

                #endregion

                TableRange = ws.Range(ws.Cell(7, 2), ws.Cell(r, 20));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range(ws.Cell(8, 4), ws.Cell(r, 4)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(8, 7), ws.Cell(r, 7)).Style.NumberFormat.Format = "0";
                ws.Range(ws.Cell(8, 8), ws.Cell(r, 20)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LostRevenueDetailReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        public ActionResult LostRevIndex(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module with CompanyId = " + CompanyId);
                DataSet ds = GetOccupancyAllFilters(CompanyId);
                var pg = getPG(ds.Tables[0]);
                var pro = getProperty(ds.Tables[1]);
                var unit = getUnit(ds.Tables[2]);
                ViewBag.propertygrp = pg;
                ViewBag.property = pro;
                ViewBag.unit = unit;
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        public ActionResult LostRevReport2(LostRevFilter obj)
        {
            try
            {
                LostRevCls _cls = new LostRevCls();
                LostRevFilter _filter = new LostRevFilter();
                _cls._filter = obj;
                int CompanyId = Convert.ToInt32(obj.CompanyId);
                string retrievequery = string.Format($@"exec LostRevenueSummaryReport @PropertyGrps='{obj.PropertyGrp}', @Proeprtys='{obj.Property}', @UnitIds='{obj.Unit}',@SDate='{obj.FromDate}',@EDate='{obj.ToDate}'");
                DBClass.SetLog("ExpenseReport retrievequery = " + retrievequery);
                DataSet ds1 = DBClass.GetData(retrievequery, CompanyId, ref errors1);
                if (ds1 != null)
                {
                    DBClass.SetLog("Getting Report View. ExpenseReport DataSet is not null");
                    if (ds1.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. ExpenseReport DataSet Table count >0 ");
                        List<LostRevList> listobj = new List<LostRevList>();
                        if (ds1.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow dr1 in ds1.Tables[0].Rows)
                            {
                                listobj.Add(new LostRevList
                                {
                                    Property = dr1["PName"].ToString(),
                                    Unit = dr1["UnitName"].ToString(),
                                    AnnualRent = Convert.ToDecimal(dr1["AnnualRent"].ToString()),
                                    FromDate = dr1["Fromdate"].ToString(),
                                    ToDate = dr1["Enddate"].ToString(),
                                    TotLostRev = Convert.ToDecimal(dr1["rentloss"].ToString()),
                                    TotLostDays = Convert.ToInt32(dr1["vaccantdays"].ToString()),
                                });
                            }

                            _cls._list = listobj;
                            Session["LostRevData"] = _cls;
                            if (_cls == null)
                            {
                                return Json("No Data", JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                return Json("Success", JsonRequestBehavior.AllowGet);
                            }
                            DBClass.SetLog("Getting Report View. ExpenseReportData body data is ready");
                        }
                        else
                        {
                            return Json("No Data", JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        return Json("No Data", JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    return Json(errors1, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult LostRevReport()
        {
            LostRevCls _data = (LostRevCls)Session["LostRevData"];
            return View(_data);
        }
        public FileResult LostRevExcel()
        {
            LostRevCls _data = (LostRevCls)Session["LostRevData"];
            LostRevFilter _head = new LostRevFilter();
            _head = _data._filter;
            List<LostRevList> _list = new List<LostRevList>();
            _list = _data._list;
            System.Data.DataTable data = new System.Data.DataTable("Lost Revenue Summary Report");
            #region DataColumns
            data.Columns.Add("Property Name", typeof(string));
            data.Columns.Add("Unit No", typeof(string));
            data.Columns.Add("Annual Rent", typeof(decimal));
            data.Columns.Add("Total Lost Days", typeof(int));
            data.Columns.Add("Total Revenue Lost", typeof(decimal));
            #endregion

            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Lost Revenue Summary Report");
                var dataTable = data;

                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 6));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Lost Revenue Summary Report";

                int r = 3;
                int cell = 2;
                ws.Cell(r, 2).Value = "From";
                ws.Cell(r, 3).Value = _list[0].FromDate;
                ws.Cell(r, 5).Value = "To";
                ws.Cell(r, 6).Value = _list[0].ToDate;

                r = 4;
                ws.Cell(r, 2).Value = "Property Group";
                ws.Cell(r, 3).Value = _head.PropertyGrpName;
                ws.Cell(r, 5).Value = "Property";
                ws.Cell(r, 6).Value = _head.PropertyName;

                r = 5;
                ws.Cell(r, 2).Value = "Unit";
                ws.Cell(r, 3).Value = _head.UnitName;
                ws.Cell(r, 5).Value = "";
                ws.Cell(r, 6).Value = "";

                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 6));
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                r = 7;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "Property Name";
                    ws.Cell(r, cell++).Value = "Unit No";
                    ws.Cell(r, cell++).Value = "Annual Rent";
                    ws.Cell(r, cell++).Value = "Total Lost Days";
                    ws.Cell(r, cell++).Value = "Total Revenue Lost";
                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r, 6));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int c = 2;

                #region TableLoop
                foreach (var obj in _list)
                {
                    c = 2;
                    r++;
                    ws.Cell(r, c++).Value = obj.Property;
                    ws.Cell(r, c++).Value = obj.Unit;
                    ws.Cell(r, c++).Value = obj.AnnualRent.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.TotLostDays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.TotLostRev.ToString("N", new CultureInfo("en-US"));
                }

                #endregion

                TableRange = ws.Range(ws.Cell(7, 2), ws.Cell(r, 6));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range(ws.Cell(8, 4), ws.Cell(r, 4)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(8, 5), ws.Cell(r, 5)).Style.NumberFormat.Format = "0";
                ws.Range(ws.Cell(8, 6), ws.Cell(r, 6)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LostRevenueSummaryReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        public ActionResult LeaseRevIndex(int CompanyId)
        {
            try
            {
                ViewBag.CompId = CompanyId;
                DBClass.SetLog("Entered to external Module with CompanyId = " + CompanyId);
                DataSet ds = GetOccupancyAllFilters(CompanyId);
                var pg = getPG(ds.Tables[0]);
                var pro = getProperty(ds.Tables[1]);
                var unit = getUnit(ds.Tables[2]);
                ViewBag.propertygrp = pg;
                ViewBag.property = pro;
                ViewBag.unit = unit;
                return View();
            }
            catch (Exception ex)
            {
                DBClass.SetLog("INdex. Exception = " + ex.Message);
                return null;
            }
        }
        public ActionResult LeaseRevReport2(LostRevFilter obj)
        {
            try
            {
                LeaseRevCls _cls = new LeaseRevCls();
                LostRevFilter _filter = new LostRevFilter();
                _cls._filter = obj;
                int CompanyId = Convert.ToInt32(obj.CompanyId);
                string retrievequery = string.Format($@"exec LeaseRentalSummaryReport @PropertyGrps='{obj.PropertyGrp}', @Proeprtys='{obj.Property}', @UnitIds='{obj.Unit}',@SDate='{obj.FromDate}',@EDate='{obj.ToDate}'");
                DBClass.SetLog("LeaseRevReport retrievequery = " + retrievequery);
                DataSet ds1 = DBClass.GetData(retrievequery, CompanyId, ref errors1);
                if (ds1 != null)
                {
                    DBClass.SetLog("Getting Report View. LeaseRevReport DataSet is not null");
                    if (ds1.Tables.Count > 0)
                    {
                        DBClass.SetLog("Getting Report View. LeaseRevReport DataSet Table count >0 ");
                        List<LeaseRevList> listobj = new List<LeaseRevList>();
                        if (ds1.Tables[0].Rows.Count > 0)
                        {
                            _cls._filter.FromDate = ds1.Tables[0].Rows[0]["Fromdate"].ToString();
                            _cls._filter.FromDate = ds1.Tables[0].Rows[0]["Todate"].ToString();
                            foreach (DataRow dr1 in ds1.Tables[0].Rows)
                            {
                                listobj.Add(new LeaseRevList
                                {
                                    Property = dr1["Property"].ToString(),
                                    Unit = dr1["Unit"].ToString(),
                                    TCNo = dr1["tcno"].ToString(),
                                    UnitType = dr1["UnitSubType"].ToString(),
                                    Usage = dr1["UsageType"].ToString(),
                                    ContractValue = Convert.ToDecimal(dr1["ContractValue"].ToString()),
                                    Status = dr1["ContractStatus"].ToString(),
                                    Tenant = dr1["Tenant"].ToString(),
                                    StartDate = dr1["StartDate2"].ToString(),
                                    EndDate = dr1["EndDate2"].ToString(),
                                    PostedDate = dr1["PostedDate2"].ToString(),
                                    TerminationDate = dr1["TerminationDate2"].ToString(),
                                    TotContractDays = Convert.ToInt32(dr1["ContractDays"].ToString()),
                                    DayRent = Convert.ToDecimal(dr1["DayRent"].ToString()),
                                    TotRevDays = Convert.ToInt32(dr1["TotalRevenueCalculatedDays"].ToString()),
                                    Am_Amt = Convert.ToDecimal(dr1["AmortizedValue"].ToString()),
                                    DeferredVal = Convert.ToDecimal(dr1["DeferredIncomeBalance"].ToString()),
                                    PDC = Convert.ToDecimal(dr1["PDCUncleared"].ToString()),
                                    Security = Convert.ToDecimal(dr1["SecurityDeposit"].ToString()),
                                    OtherIncome = Convert.ToDecimal(dr1["OtherIncome"].ToString()),
                                });
                            }

                            _cls._list = listobj;
                            Session["LeaseRevData"] = _cls;
                            if (_cls == null)
                            {
                                return Json("No Data", JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                return Json("Success", JsonRequestBehavior.AllowGet);
                            }
                            DBClass.SetLog("Getting Report View. LeaseRevReportData body data is ready");
                        }
                        else
                        {
                            return Json("No Data", JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        return Json("No Data", JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    return Json(errors1, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                DBClass.SetLog("Getting Report View. EXCEPTION = " + ex.Message);
                return null;
            }
        }
        public ActionResult LeaseRevReport()
        {
            LeaseRevCls _data = (LeaseRevCls)Session["LeaseRevData"];
            return View(_data);
        }
        public FileResult LeaseRevExcel()
        {
            LeaseRevCls _data = (LeaseRevCls)Session["LeaseRevData"];
            LostRevFilter _head = new LostRevFilter();
            _head = _data._filter;
            List<LeaseRevList> _list = new List<LeaseRevList>();
            _list = _data._list;
            System.Data.DataTable data = new System.Data.DataTable("Lease Rental Summary Report");
            #region DataColumns
            data.Columns.Add("Property", typeof(string));
            data.Columns.Add("Unit", typeof(string));
            data.Columns.Add("TCNo", typeof(string));
            data.Columns.Add("UnitType", typeof(string));
            data.Columns.Add("Usage", typeof(string));
            data.Columns.Add("ContractValue", typeof(decimal));
            data.Columns.Add("Status", typeof(string));
            data.Columns.Add("Tenant", typeof(string));
            data.Columns.Add("StartDate", typeof(string));
            data.Columns.Add("EndDate", typeof(string));
            data.Columns.Add("PostedDate", typeof(string));
            data.Columns.Add("TerminationDate", typeof(string));
            data.Columns.Add("TotContractDays", typeof(int));
            data.Columns.Add("DayRent", typeof(decimal));
            data.Columns.Add("TotRevDays", typeof(int));
            data.Columns.Add("Am_Amt", typeof(decimal));
            data.Columns.Add("DeferredVal", typeof(decimal));
            data.Columns.Add("PDC", typeof(decimal));
            data.Columns.Add("Security", typeof(decimal));
            data.Columns.Add("OtherIncome", typeof(decimal));
            #endregion

            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Lease Rental Summary Report");
                var dataTable = data;

                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 21));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "Lease Rental Summary Report";

                int r = 3;
                int cell = 2;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 5)).Merge().Value = "From";
                ws.Range(ws.Cell(r, 6), ws.Cell(r, 10)).Merge().Value = _head.FromDate;
                ws.Range(ws.Cell(r, 13), ws.Cell(r, 16)).Merge().Value = "To";
                ws.Range(ws.Cell(r, 17), ws.Cell(r, 21)).Merge().Value = _head.ToDate;

                r = 4;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 5)).Merge().Value = "Property Group";
                ws.Range(ws.Cell(r, 6), ws.Cell(r, 10)).Merge().Value = _head.PropertyGrpName;
                ws.Range(ws.Cell(r, 13), ws.Cell(r, 16)).Merge().Value = "Property";
                ws.Range(ws.Cell(r, 17), ws.Cell(r, 21)).Merge().Value = _head.PropertyName;

                r = 5;
                ws.Range(ws.Cell(r, 2), ws.Cell(r, 5)).Merge().Value = "Unit";
                ws.Range(ws.Cell(r, 6), ws.Cell(r, 10)).Merge().Value = _head.UnitName;
                ws.Range(ws.Cell(r, 13), ws.Cell(r, 16)).Merge().Value = "";
                ws.Range(ws.Cell(r, 17), ws.Cell(r, 21)).Merge().Value = "";

                var TableRange = ws.Range(ws.Cell(3, 2), ws.Cell(r, 21));
                TableRange.Style.Fill.BackgroundColor = XLColor.White;
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(ws.Cell(3, 2), ws.Cell(5, 5)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Range(ws.Cell(3, 13), ws.Cell(5, 16)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                ws.Range(ws.Cell(3, 6), ws.Cell(5, 10)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(ws.Cell(3, 17), ws.Cell(5, 21)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                r = 7;
                cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(r, cell++).Value = "TC No";
                    ws.Cell(r, cell++).Value = "Status";
                    ws.Cell(r, cell++).Value = "Tenant";
                    ws.Cell(r, cell++).Value = "Property";
                    ws.Cell(r, cell++).Value = "Unit";
                    ws.Cell(r, cell++).Value = "Usage Type";
                    ws.Cell(r, cell++).Value = "Unit Type";
                    ws.Cell(r, cell++).Value = "Contract Value";
                    ws.Cell(r, cell++).Value = "Start Date";
                    ws.Cell(r, cell++).Value = "End Date";
                    ws.Cell(r, cell++).Value = "Revenue Posted Date";
                    ws.Cell(r, cell++).Value = "Total no of Contract";
                    ws.Cell(r, cell++).Value = "Rent Per Day";
                    ws.Cell(r, cell++).Value = "Total Revenue Calculated Days";   
                    ws.Cell(r, cell++).Value = "Amortized Value";
                    ws.Cell(r, cell++).Value = "Deferred Income Balance";
                    ws.Cell(r, cell++).Value = "PDC Uncleared";
                    ws.Cell(r, cell++).Value = "Security Deposit";
                    ws.Cell(r, cell++).Value = "Other Income";
                    ws.Cell(r, cell++).Value = "Termination Date";
                    #endregion
                }
                TableRange = ws.Range(ws.Cell(r, 2), ws.Cell(r, 21));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int c = 2;

                #region TableLoop
                foreach (var obj in _list)
                {
                    c = 2;
                    r++;
                    ws.Cell(r, c++).Value = obj.TCNo;
                    ws.Cell(r, c++).Value = obj.Status;
                    ws.Cell(r, c++).Value = obj.Tenant;
                    ws.Cell(r, c++).Value = obj.Property;
                    ws.Cell(r, c++).Value = obj.Unit;
                    ws.Cell(r, c++).Value = obj.Usage;
                    ws.Cell(r, c++).Value = obj.UnitType;
                    ws.Cell(r, c++).Value = obj.ContractValue.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.StartDate;
                    ws.Cell(r, c++).Value = obj.EndDate;
                    ws.Cell(r, c++).Value = obj.PostedDate;
                    ws.Cell(r, c++).Value = obj.TotContractDays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.DayRent.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.TotRevDays.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Am_Amt.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.DeferredVal.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.PDC.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.Security.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.OtherIncome.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = obj.TerminationDate;
                }

                #endregion

                TableRange = ws.Range(ws.Cell(7, 2), ws.Cell(r, 21));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range(ws.Cell(8, 9), ws.Cell(r, 9)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(8, 13), ws.Cell(r, 13)).Style.NumberFormat.Format = "0";
                ws.Range(ws.Cell(8, 15), ws.Cell(r, 15)).Style.NumberFormat.Format = "0";
                ws.Range(ws.Cell(8, 14), ws.Cell(r, 14)).Style.NumberFormat.Format = "0.00";
                ws.Range(ws.Cell(8, 16), ws.Cell(r, 20)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LeaseRentalSummaryReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
    }
}